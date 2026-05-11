"""Bus-Level Inertia Assessment Algorithm.

Implements the Ghosh, Isbeih & El Moursi methodology:
  "Bus-Level Inertia Assessment for Converter-Dominated Power Systems"
  IEEE Transactions on Power Delivery, Vol. 38, No. 4, Aug. 2023.

Core equations (from paper):
  Eq. 14:  H_B(t) = W_c(t) x H_G
  Eq. 15:  W_c(t) = { -Y_bb^{-1}(t) x Y_bg(t) ./ v_b(t) } .* e_g(t)
  Eq. 16:  Sys_str(t) = sum( H_B(t) )

Where:
  Y_bb  = admittance submatrix among load buses (nb x nb)
  Y_bg  = admittance submatrix from load buses to generator buses (nb x ng)
  v_b   = complex bus voltage magnitudes at load buses (nb,)
  e_g   = binary generator participation vector (ng,) [1 if gen online, 0 if tripped]
  H_G   = generator inertia constants (ng,) [seconds]
  W_c   = inertia weighting matrix (nb x ng)
  H_B   = bus-level inertia (nb,) for load buses [seconds]
"""

import numpy as np
from core.ybus import build_ybus, partition_ybus


def compute_weighting_matrix(Y_bb, Y_bg, v_b, e_g, fault=False):
    """Compute the inertia weighting matrix W_c (Eq. 15).

    Parameters
    ----------
    Y_bb  : complex ndarray (nb x nb)  load-bus admittance submatrix
    Y_bg  : complex ndarray (nb x ng)  load-to-gen admittance submatrix
    v_b   : real ndarray (nb,)          bus voltage magnitudes at load buses [pu]
    e_g   : real ndarray (ng,)          generator participation (0 or 1)
    fault : bool
        If True, assume v_b = 1.0 pu for all load buses (fault condition
        where voltages are not measurable), per paper Section IV.

    Returns
    -------
    W_c : real ndarray (nb x ng)
        Inertia weighting matrix. Rows sum to 1 when all generators
        are online; row sums < 1 when generators are tripped (H_B
        decreases proportionally with lost inertia sources).
    """
    Y_bb_inv = np.linalg.inv(Y_bb)

    # Core term: -Y_bb^{-1} * Y_bg   shape (nb x ng)
    M = -Y_bb_inv @ Y_bg

    # Take magnitude of each element (W_c is real-valued per paper)
    M_mag = np.abs(M)

    if fault:
        v_norm = np.ones(Y_bb.shape[0])
    else:
        v_norm = np.where(np.abs(v_b) > 0.01, np.abs(v_b), 1.0)

    # Element-wise division by v_b (broadcasting: nb x ng ./ nb x 1)
    W_raw = M_mag / v_norm[:, np.newaxis]

    # Normalise using FULL matrix (all generators online) so rows sum to 1
    # at baseline — physical meaning: H_B = weighted average of all H_G
    row_sums_full = W_raw.sum(axis=1, keepdims=True)
    row_sums_full = np.where(row_sums_full < 1e-12, 1.0, row_sums_full)
    W_normalised = W_raw / row_sums_full

    # Apply e_g AFTER normalisation: tripped generators → zero contribution
    # Row sums become < 1, so H_B correctly decreases with generator loss
    W_c = W_normalised * e_g[np.newaxis, :]

    return W_c


def compute_bus_inertia(system_data, v_bus=None, gen_status=None):
    """Compute bus-level inertia for all buses in the system.

    Parameters
    ----------
    system_data : module
        IEEE data module (ieee14, ieee39, etc.) with:
        N_BUS, GEN_BUSES, H_GEN, BRANCH_DATA attributes.
    v_bus : ndarray (n_bus,) or None
        Bus voltage magnitudes in pu. If None, assume 1.0 pu everywhere.
    gen_status : ndarray (n_gen,) or None
        Generator participation flags (1=online, 0=tripped).
        If None, all generators are online.

    Returns
    -------
    H_B_all : ndarray (n_bus,)
        Bus-level inertia for every bus [seconds].
        Generator buses carry their own H_G value.
    W_c : ndarray (nb x ng)
        Weighting matrix (load buses only).
    gen_idx  : list[int]  0-indexed generator bus positions
    load_idx : list[int]  0-indexed load bus positions
    """
    n = system_data.N_BUS
    H_G = system_data.H_GEN
    gen_buses = system_data.GEN_BUSES

    if v_bus is None:
        v_bus = np.ones(n)
    if gen_status is None:
        gen_status = np.ones(len(H_G))

    # Build Y-bus
    Y = build_ybus(n, system_data.BRANCH_DATA)

    # Partition
    Y_gg, Y_gb, Y_bg, Y_bb, gen_idx, load_idx = partition_ybus(Y, gen_buses)

    # Extract load-bus voltages
    v_b = v_bus[load_idx]

    # Compute weighting matrix
    W_c = compute_weighting_matrix(Y_bb, Y_bg, v_b, gen_status)

    # Bus inertia for load buses: H_B = W_c @ H_G  (nb,)
    H_B_load = W_c @ (H_G * gen_status)

    # Assemble full vector (all buses)
    H_B_all = np.zeros(n)
    for k, li in enumerate(load_idx):
        H_B_all[li] = H_B_load[k]
    for k, gi in enumerate(gen_idx):
        H_B_all[gi] = H_G[k] * gen_status[k]

    return H_B_all, W_c, gen_idx, load_idx


def system_strength(H_B_all):
    """Total system inertia (Eq. 16): Sys_str = sum(H_B)."""
    return H_B_all.sum()


def run_res_penetration_study(system_data, res_levels=None):
    """Study effect of increasing RES penetration on bus inertia.

    Simulates progressive generator displacement by RES (wind/solar):
    generators are tripped in order of increasing inertia constant
    to represent high-RES scenarios.

    Parameters
    ----------
    system_data : module
        IEEE data module.
    res_levels : list[float]
        RES penetration levels as fractions of generators displaced.
        Default: [0.0, 0.1, 0.2, ..., 0.9]

    Returns
    -------
    results : list of dict
        Each dict: {'res_pct', 'sys_str', 'H_B', 'gen_status'}
    """
    if res_levels is None:
        res_levels = np.arange(0.0, 1.0, 0.1)

    H_G = system_data.H_GEN
    n_gen = len(H_G)

    # Trip order: generators with smallest H first (most easily replaced)
    trip_order = np.argsort(H_G)
    results = []

    for pct in res_levels:
        n_trip = int(round(pct * n_gen))
        gen_status = np.ones(n_gen)
        if n_trip > 0:
            gen_status[trip_order[:n_trip]] = 0.0

        H_B, W_c, gi, li = compute_bus_inertia(system_data, gen_status=gen_status)
        sys_str = system_strength(H_B)

        results.append({
            'res_pct':    pct * 100,
            'sys_str':    sys_str,
            'H_B':        H_B,
            'gen_status': gen_status,
            'n_online':   int(gen_status.sum()),
        })

    return results


def contingency_analysis(system_data, n_contingencies=3):
    """N-k contingency analysis on generator trips.

    Parameters
    ----------
    system_data : module
    n_contingencies : int
        Maximum k for N-k analysis.

    Returns
    -------
    contingencies : list of dict
        Sorted by descending impact on system strength.
    """
    import itertools
    H_G = system_data.H_GEN
    n_gen = len(H_G)

    # Baseline
    H_B0, _, _, _ = compute_bus_inertia(system_data)
    S0 = system_strength(H_B0)

    contingencies = []
    for k in range(1, n_contingencies + 1):
        for combo in itertools.combinations(range(n_gen), k):
            gen_status = np.ones(n_gen)
            gen_status[list(combo)] = 0.0
            H_B, _, _, _ = compute_bus_inertia(system_data, gen_status=gen_status)
            S = system_strength(H_B)
            contingencies.append({
                'tripped_gens': combo,
                'tripped_buses': [system_data.GEN_BUSES[i] for i in combo],
                'sys_str':      S,
                'delta_str':    S0 - S,
                'pct_loss':     (S0 - S) / S0 * 100,
            })

    contingencies.sort(key=lambda x: -x['delta_str'])
    return contingencies
