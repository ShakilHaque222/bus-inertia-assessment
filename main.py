"""Bus-Level Inertia Assessment – Main Runner.

Implements Ghosh, Isbeih & El Moursi (IEEE Trans. Power Del., 2023)
for IEEE 14/39/68/118-bus test systems.

Usage:
    python main.py                    # run all systems
    python main.py --system 39        # run only IEEE 39-bus
    python main.py --system 39 --no-freq   # skip freq simulation
"""

import os
import sys
import argparse
import numpy as np

# ── Ensure project root is on path ───────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))

import data.ieee14  as d14
import data.ieee39  as d39
import data.ieee68  as d68
import data.ieee118 as d118

from core.bus_inertia import (
    compute_bus_inertia, system_strength,
    run_res_penetration_study, contingency_analysis
)
from core.freq_sim import simulate_all_buses
from viz.plots import (
    plot_network_topology, plot_bus_inertia_bar,
    plot_frequency_response, plot_weighting_matrix,
    plot_res_impact, plot_system_comparison,
    POS_14, POS_39, OUT_DIR
)

DIVIDER = '=' * 60


def print_header(title):
    print(f'\n{DIVIDER}')
    print(f'  {title}')
    print(DIVIDER)


def print_table(rows, col_widths, headers):
    fmt = '  '.join(f'{{:<{w}}}' for w in col_widths)
    sep = '  '.join('-' * w for w in col_widths)
    print(fmt.format(*headers))
    print(sep)
    for r in rows:
        print(fmt.format(*[str(v) for v in r]))


def run_system(sdata, name, short_name, pos=None, seed=42,
               run_freq=True, run_res=True, run_contingency=False):
    """Full pipeline for one IEEE test system."""
    print_header(f'{name}  ({sdata.N_BUS} buses, {sdata.N_GEN} generators)')

    # ── 1. Compute bus inertia (baseline, all generators online) ─────────────
    print('\n[1] Computing bus-level inertia...')
    H_B_all, W_c, gen_idx, load_idx = compute_bus_inertia(sdata)
    sys_str = system_strength(H_B_all)

    # Print top-10 and bottom-10
    sorted_buses = np.argsort(H_B_all)
    print(f'\n  System Strength (sum H_B) = {sys_str:.2f} s')
    print(f'  Avg bus inertia           = {H_B_all.mean():.2f} s')
    print(f'\n  Top-5 high-inertia buses:')
    rows_hi = [(f'Bus {sorted_buses[-k]+1}',
                f'{H_B_all[sorted_buses[-k]]:.2f} s',
                'GEN' if (sorted_buses[-k]+1) in set(sdata.GEN_BUSES) else 'LOAD')
               for k in range(1, 6)]
    print_table(rows_hi, [12, 14, 6], ['Bus', 'H_B', 'Type'])

    print(f'\n  Bottom-5 low-inertia buses:')
    rows_lo = [(f'Bus {sorted_buses[k]+1}',
                f'{H_B_all[sorted_buses[k]]:.2f} s',
                'GEN' if (sorted_buses[k]+1) in set(sdata.GEN_BUSES) else 'LOAD')
               for k in range(5)]
    print_table(rows_lo, [12, 14, 6], ['Bus', 'H_B', 'Type'])

    # ── 2. Network topology plot ──────────────────────────────────────────────
    print(f'\n[2] Plotting network topology...')
    plot_network_topology(sdata, H_B_all, name, pos=pos, seed=seed)

    # ── 3. Bar chart ──────────────────────────────────────────────────────────
    print(f'[3] Plotting inertia bar chart...')
    plot_bus_inertia_bar(sdata, H_B_all, name)

    # ── 4. Weighting matrix ───────────────────────────────────────────────────
    print(f'[4] Plotting weighting matrix...')
    plot_weighting_matrix(W_c, sdata, name)

    # ── 5. Frequency response ─────────────────────────────────────────────────
    if run_freq:
        print(f'[5] Simulating frequency response (delta_P = 0.1 pu)...')
        t, freq_all = simulate_all_buses(sdata, H_B_all, delta_P=0.1,
                                          D=0.01, f0=60.0, t_end=10.0)
        plot_frequency_response(t, freq_all, sdata, H_B_all, name)
        print(f'    Freq simulation done.')

        # Print nadir comparison for high vs low inertia buses
        load_mask = [i for i in range(sdata.N_BUS)
                     if (i+1) not in set(sdata.GEN_BUSES)]
        H_load = H_B_all[load_mask]
        i_max = load_mask[np.argmax(H_load)]
        i_min = load_mask[np.argmin(H_load)]
        f_nadir_max = freq_all[i_max].min()
        f_nadir_min = freq_all[i_min].min()
        print(f'\n  Freq nadir comparison:')
        print(f'    High-inertia Bus {i_max+1} (H_B={H_B_all[i_max]:.1f}s): '
              f'nadir = {f_nadir_max:.4f} Hz')
        print(f'    Low-inertia  Bus {i_min+1} (H_B={H_B_all[i_min]:.1f}s): '
              f'nadir = {f_nadir_min:.4f} Hz')
    else:
        t, freq_all = None, None

    # ── 6. RES penetration study ──────────────────────────────────────────────
    if run_res:
        print(f'[6] Running RES penetration study...')
        res_results = run_res_penetration_study(sdata)
        plot_res_impact(res_results, name)

        print(f'  RES penetration impact on system strength:')
        rows_res = [(f"{r['res_pct']:.0f}%",
                     f"{r['sys_str']:.1f} s",
                     f"{r['n_online']}/{sdata.N_GEN}")
                    for r in res_results[::2]]
        print_table(rows_res, [8, 14, 10], ['RES%', 'Sys.Strength', 'Gens On'])

    # ── 7. Contingency analysis (optional, slow for large systems) ────────────
    if run_contingency and sdata.N_GEN <= 16:
        print(f'[7] Running N-1 contingency analysis...')
        contingencies = contingency_analysis(sdata, n_contingencies=1)
        print(f'  Top-5 most critical generator outages:')
        rows_c = [(f"Bus {c['tripped_buses'][0]}",
                   f"{c['sys_str']:.1f} s",
                   f"{c['pct_loss']:.1f}%")
                  for c in contingencies[:5]]
        print_table(rows_c, [12, 14, 10], ['Tripped Bus', 'Sys.Strength', 'Loss%'])

    return {
        'name':    name,
        'n_bus':   sdata.N_BUS,
        'n_gen':   sdata.N_GEN,
        'H_B':     H_B_all,
        'W_c':     W_c,
        'sys_str': sys_str,
    }


def main():
    parser = argparse.ArgumentParser(description='Bus-Level Inertia Assessment')
    parser.add_argument('--system', type=str, default='all',
                        choices=['14', '39', '68', '118', 'all'],
                        help='IEEE system to analyse (default: all)')
    parser.add_argument('--no-freq', action='store_true',
                        help='Skip frequency response simulation')
    parser.add_argument('--no-res', action='store_true',
                        help='Skip RES penetration study')
    parser.add_argument('--contingency', action='store_true',
                        help='Run N-1 contingency analysis')
    args = parser.parse_args()

    os.makedirs(OUT_DIR, exist_ok=True)
    run_freq = not args.no_freq
    run_res  = not args.no_res

    systems_to_run = {
        '14':  (d14,  'IEEE 14-Bus',  '14',  POS_14, 42),
        '39':  (d39,  'IEEE 39-Bus',  '39',  POS_39, 42),
        '68':  (d68,  'IEEE 68-Bus',  '68',  None,   7),
        '118': (d118, 'IEEE 118-Bus', '118', None,   13),
    }

    if args.system != 'all':
        systems_to_run = {args.system: systems_to_run[args.system]}

    all_results = []
    for key, (sdata, name, short, pos, seed) in systems_to_run.items():
        result = run_system(
            sdata, name, short, pos=pos, seed=seed,
            run_freq=run_freq, run_res=run_res,
            run_contingency=args.contingency
        )
        all_results.append(result)

    # ── Summary comparison plot ───────────────────────────────────────────────
    if len(all_results) > 1:
        print_header('Generating system comparison chart')
        plot_system_comparison(all_results)

    # ── Print final summary table ─────────────────────────────────────────────
    print_header('SUMMARY – All IEEE Test Systems')
    rows = [
        (r['name'], r['n_bus'], r['n_gen'],
         f"{r['sys_str']:.1f} s",
         f"{r['H_B'].mean():.2f} s",
         f"{r['H_B'].max():.1f} s (Bus {np.argmax(r['H_B'])+1})",
         f"{r['H_B'].min():.1f} s (Bus {np.argmin(r['H_B'])+1})")
        for r in all_results
    ]
    print_table(rows,
                [14, 6, 6, 14, 12, 22, 22],
                ['System', 'Buses', 'Gens', 'Sys.Strength',
                 'Avg H_B', 'Max H_B', 'Min H_B'])

    print(f'\nAll figures saved to: {OUT_DIR}')
    print(DIVIDER)


if __name__ == '__main__':
    main()
