"""Frequency response simulation using the swing equation.

Simulates per-bus frequency deviation after a generation loss event,
using bus-level inertia H_B as the effective inertia at each bus.

Swing equation (per unit, per bus):
    2*H_B / f0 * df/dt = Pm - Pe - D*delta_f

For simplified post-fault analysis:
    df/dt = -f0/(2*H_B) * (delta_P + D*delta_f)

Where:
  H_B    = bus-level inertia [s]
  f0     = nominal frequency [Hz] (50 or 60 Hz)
  delta_P = per-unit power imbalance at the bus
  D      = load damping coefficient [pu/pu]
  ROCOF  = -f0 * delta_P / (2 * H_B)  [Hz/s]
"""

import numpy as np
from scipy.integrate import solve_ivp


def rocof(H_B, delta_P, f0=60.0):
    """Compute Rate of Change of Frequency (ROCOF) at each bus.

    ROCOF = -f0 * delta_P / (2 * H_B)   [Hz/s]

    Parameters
    ----------
    H_B     : ndarray (n_bus,)  bus inertia [s]
    delta_P : float             power imbalance [pu] (positive = generation loss)
    f0      : float             nominal frequency [Hz]

    Returns
    -------
    rocof_vals : ndarray (n_bus,)  [Hz/s] — negative means frequency drops
    """
    H_safe = np.where(H_B > 0.01, H_B, np.inf)
    return -f0 * delta_P / (2.0 * H_safe)


def simulate_frequency_response(H_B_bus, bus_idx, delta_P=0.1,
                                 D=0.01, f0=60.0,
                                 t_end=10.0, n_points=1000,
                                 droop_R=0.05, droop_delay=0.5):
    """Simulate frequency deviation at a specific bus after generation loss.

    Uses a simplified single-bus equivalent with primary frequency regulation.

    Parameters
    ----------
    H_B_bus  : float   bus-level inertia constant [s]
    bus_idx  : int     bus number (for labelling)
    delta_P  : float   per-unit generation loss [pu on system base]
    D        : float   load damping coefficient [pu/Hz]
    f0       : float   nominal frequency [Hz]
    t_end    : float   simulation end time [s]
    n_points : int     number of output time steps
    droop_R  : float   governor droop setting [pu]
    droop_delay : float governor response delay [s]

    Returns
    -------
    t  : ndarray (n_points,)   time [s]
    df : ndarray (n_points,)   frequency deviation [Hz]
    """
    if H_B_bus < 0.1:
        H_B_bus = 0.1

    def swing_eq(t, y):
        df = y[0]   # frequency deviation [Hz]
        # Governor action (simplified first-order with delay)
        Pg = 0.0
        if t > droop_delay:
            Pg = -df / (droop_R * f0)  # governor responds to freq drop

        ddf = (f0 / (2.0 * H_B_bus)) * (-delta_P - D * df + Pg)
        return [ddf]

    t_span = (0.0, t_end)
    t_eval = np.linspace(0, t_end, n_points)

    sol = solve_ivp(swing_eq, t_span, [0.0], t_eval=t_eval,
                    method='RK45', rtol=1e-6, atol=1e-8)
    return sol.t, sol.y[0]


def simulate_all_buses(system_data, H_B_all, delta_P=0.1,
                       D=0.01, f0=60.0, t_end=10.0):
    """Simulate frequency response for all buses in the system.

    Parameters
    ----------
    system_data : module   IEEE data module
    H_B_all     : ndarray (n_bus,)  bus-level inertia
    delta_P     : float   per-unit generation loss
    D           : float   load damping coefficient
    f0          : float   nominal frequency [Hz]
    t_end       : float   simulation duration [s]

    Returns
    -------
    t       : ndarray (n_points,)        time vector [s]
    freq_all: ndarray (n_bus, n_points)  frequency deviation at each bus [Hz]
    """
    n_bus = system_data.N_BUS
    n_pts = 500
    t = np.linspace(0, t_end, n_pts)
    freq_all = np.zeros((n_bus, n_pts))

    for i in range(n_bus):
        _, df = simulate_frequency_response(
            H_B_all[i], i + 1, delta_P=delta_P,
            D=D, f0=f0, t_end=t_end, n_points=n_pts
        )
        freq_all[i, :] = f0 + df

    return t, freq_all


def nadir_frequency(t, freq):
    """Return the nadir (minimum) frequency and time of occurrence."""
    idx = np.argmin(freq)
    return freq[idx], t[idx]


def frequency_settled(t, freq, f0=60.0, tol=0.01, window=2.0):
    """Estimate steady-state settled frequency after the transient."""
    dt = t[1] - t[0]
    n_window = int(window / dt)
    if n_window < 2:
        return freq[-1]
    tail = freq[-n_window:]
    if np.ptp(tail) < tol:
        return np.mean(tail)
    return freq[-1]
