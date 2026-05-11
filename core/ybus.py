"""Y-bus matrix construction and partitioning.

Builds the admittance matrix from branch data and partitions it into
generator (g) and load (b) bus submatrices as required by the
Ghosh et al. bus-level inertia algorithm (IEEE Trans. Power Del., 2023).
"""

import numpy as np


def build_ybus(n_bus, branch_data):
    """Build the full Y-bus admittance matrix from branch data.

    Parameters
    ----------
    n_bus : int
        Total number of buses (buses are 1-indexed in branch_data).
    branch_data : list of [from, to, R, X, B_half, tap]
        Branch parameters in per-unit. tap=0 means transmission line.

    Returns
    -------
    Y : complex ndarray, shape (n_bus, n_bus)
        Full admittance matrix (0-indexed).
    """
    Y = np.zeros((n_bus, n_bus), dtype=complex)

    for branch in branch_data:
        fr, to = int(branch[0]) - 1, int(branch[1]) - 1
        R, X, B_half = branch[2], branch[3], branch[4]
        tap = branch[5]

        z = complex(R, X)
        if abs(z) < 1e-12:
            z = complex(0, 1e-6)   # avoid division by zero on ideal XFMRs
        y = 1.0 / z
        b_sh = complex(0, B_half)

        if tap == 0:
            # Transmission line (pi model)
            Y[fr, fr] += y + b_sh
            Y[to, to] += y + b_sh
            Y[fr, to] -= y
            Y[to, fr] -= y
        else:
            # Transformer (off-nominal turns ratio)
            a = tap
            Y[fr, fr] += y / (a ** 2)
            Y[to, to] += y
            Y[fr, to] -= y / a
            Y[to, fr] -= y / a

    return Y


def partition_ybus(Y, gen_buses_1idx):
    """Partition Y-bus into generator (g) and load (b) bus submatrices.

    Parameters
    ----------
    Y : complex ndarray, shape (n, n)
        Full Y-bus matrix (0-indexed).
    gen_buses_1idx : list[int]
        Generator bus numbers (1-indexed).

    Returns
    -------
    Y_gg : complex ndarray  shape (ng, ng)
    Y_gb : complex ndarray  shape (ng, nb)
    Y_bg : complex ndarray  shape (nb, ng)
    Y_bb : complex ndarray  shape (nb, nb)
    gen_idx   : list[int]   0-indexed generator bus positions
    load_idx  : list[int]   0-indexed load (non-generator) bus positions
    """
    n = Y.shape[0]
    gen_idx  = [b - 1 for b in gen_buses_1idx]
    load_idx = [i for i in range(n) if i not in gen_idx]

    g = np.array(gen_idx)
    b = np.array(load_idx)

    Y_gg = Y[np.ix_(g, g)]
    Y_gb = Y[np.ix_(g, b)]
    Y_bg = Y[np.ix_(b, g)]
    Y_bb = Y[np.ix_(b, b)]

    return Y_gg, Y_gb, Y_bg, Y_bb, gen_idx, load_idx


def kron_reduce(Y, gen_buses_1idx):
    """Kron-reduce Y-bus to generator buses only (Yred).

    Yred = Y_gg - Y_gb @ inv(Y_bb) @ Y_bg

    Useful for classical generator stability analysis.
    """
    _, Y_gb, Y_bg, Y_bb, _, _ = partition_ybus(Y, gen_buses_1idx)
    Yred = Y_gg = None  # silence linter
    _, Y_gb, Y_bg, Y_bb, _, _ = partition_ybus(Y, gen_buses_1idx)
    Y_gg, *_ = partition_ybus(Y, gen_buses_1idx)
    Y_gg = Y_gg  # noqa
    Y_bb_inv = np.linalg.inv(Y_bb)
    return Y_gg - Y_gb @ Y_bb_inv @ Y_bg
