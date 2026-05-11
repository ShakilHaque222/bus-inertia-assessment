"""Visualization module for bus-level inertia results.

All plots use a dark academic theme consistent with the PowerPoint presentation.
Generates publication-quality figures for:
  - Network topology with inertia overlaid
  - Bus inertia distribution (bar/heatmap)
  - Frequency response curves
  - Weighting matrix heatmap
  - RES penetration impact
  - System strength comparison
"""

import os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.cm as cm
import networkx as nx

# ── Theme ─────────────────────────────────────────────────────────────────────
DARK = {
    'bg':     '#0d0d1f',
    'panel':  '#1a1a35',
    'gen':    '#e84040',
    'load':   '#2266ee',
    'edge':   '#5588bb',
    'text':   '#e8e8f0',
    'accent': '#ffd700',
    'green':  '#44cc88',
    'orange': '#ff8c00',
    'cyan':   '#00cccc',
}

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'results')

# ── Layout positions for IEEE 14-bus (manually placed for clarity) ────────────
POS_14 = {
    1: (0, 4),   2: (2, 6),   3: (5, 6),   4: (4, 4),   5: (2, 2),
    6: (7, 5),   7: (6, 3),   8: (8, 3),   9: (6, 1),  10: (7, 2),
   11: (8, 4),  12: (9, 6),  13: (9, 4),  14: (8, 1),
}

# ── IEEE 39-bus New England positions (approximate geographic layout) ─────────
POS_39 = {
    1: (7, 6),   2: (6, 7),   3: (5, 8),   4: (4, 7),   5: (4, 5),
    6: (3, 4),   7: (2, 5),   8: (2, 3),   9: (3, 2),  10: (4, 1),
   11: (5, 2),  12: (6, 3),  13: (6, 1),  14: (7, 2),  15: (8, 1),
   16: (9, 2),  17: (9, 4),  18: (8, 5),  19: (10, 3), 20: (11, 4),
   21: (11, 6), 22: (10, 7), 23: (9, 8),  24: (8, 7),  25: (6, 9),
   26: (5, 10), 27: (4, 9),  28: (3, 8),  29: (2, 9),
   30: (6, 11), 31: (3, 3),  32: (4, 0),  33: (10, 1), 34: (12, 5),
   35: (10, 9), 36: (9, 10), 37: (6, 12), 38: (2, 11), 39: (8, 9),
}


def _style_ax(ax, title=''):
    ax.set_facecolor(DARK['bg'])
    ax.tick_params(colors=DARK['text'], labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor(DARK['panel'])
    if title:
        ax.set_title(title, color=DARK['text'], fontsize=11, pad=8)


def _save(fig, fname):
    path = os.path.join(OUT_DIR, fname)
    fig.savefig(path, dpi=150, bbox_inches='tight',
                facecolor=DARK['bg'], edgecolor='none')
    plt.close(fig)
    print(f"  SAVED  {fname}")
    return path


# ── 1. Network topology ───────────────────────────────────────────────────────

def plot_network_topology(system_data, H_B_all, name, pos=None, seed=42):
    """Draw network with buses coloured by bus-level inertia.

    Parameters
    ----------
    system_data : module    IEEE data module
    H_B_all     : ndarray   bus-level inertia (n_bus,)
    name        : str       system name e.g. 'IEEE 14-Bus'
    pos         : dict or None  {bus_1idx: (x, y)} if pre-defined
    seed        : int       spring_layout random seed
    """
    n = system_data.N_BUS
    gen_set = set(system_data.GEN_BUSES)

    G = nx.Graph()
    for i in range(1, n + 1):
        G.add_node(i)
    for br in system_data.BRANCH_DATA:
        G.add_edge(int(br[0]), int(br[1]))

    if pos is None:
        pos = nx.spring_layout(G, seed=seed, k=2.5 / np.sqrt(n))

    # Normalise positions to [0,10] x [0,8]
    xs = np.array([pos[k][0] for k in pos])
    ys = np.array([pos[k][1] for k in pos])
    rx, ry = xs.max() - xs.min(), ys.max() - ys.min()
    if rx < 1e-6: rx = 1
    if ry < 1e-6: ry = 1
    pos_n = {k: ((pos[k][0] - xs.min()) / rx * 10,
                  (pos[k][1] - ys.min()) / ry * 8) for k in pos}

    # Figure
    fig, ax = plt.subplots(figsize=(14, 9))
    fig.patch.set_facecolor(DARK['bg'])
    ax.set_facecolor(DARK['bg'])
    ax.axis('off')

    # Edges
    for u, v in G.edges():
        x0, y0 = pos_n[u]
        x1, y1 = pos_n[v]
        ax.plot([x0, x1], [y0, y1], color=DARK['edge'], lw=0.8, alpha=0.6, zorder=1)

    # Colour map for H_B
    h_vals = H_B_all
    h_min, h_max = h_vals.min(), h_vals.max()
    if h_max - h_min < 1e-6:
        h_max = h_min + 1

    cmap = cm.get_cmap('plasma')
    norm = mcolors.Normalize(vmin=h_min, vmax=h_max)

    node_size_gen  = 320
    node_size_load = 160

    for bus in range(1, n + 1):
        x, y = pos_n[bus]
        h = H_B_all[bus - 1]
        color = cmap(norm(h))
        is_gen = bus in gen_set
        size = node_size_gen if is_gen else node_size_load
        marker = 's' if is_gen else 'o'

        ax.scatter(x, y, s=size, c=[color], marker=marker,
                   edgecolors='white', linewidths=0.6, zorder=3)

        # Label for small networks
        if n <= 39:
            ax.text(x, y + 0.28, str(bus), color='white', fontsize=7,
                    ha='center', va='bottom', zorder=4,
                    fontweight='bold' if is_gen else 'normal')

    # Colorbar
    sm = cm.ScalarMappable(cmap=cmap, norm=norm)
    sm.set_array([])
    cbar = fig.colorbar(sm, ax=ax, fraction=0.025, pad=0.02)
    cbar.set_label('Bus Inertia H_B [s]', color=DARK['text'], fontsize=10)
    cbar.ax.yaxis.set_tick_params(color=DARK['text'])
    plt.setp(cbar.ax.yaxis.get_ticklabels(), color=DARK['text'])

    # Legend
    from matplotlib.lines import Line2D
    legend_els = [
        Line2D([0], [0], marker='s', color='w', markerfacecolor=DARK['gen'],
               markersize=9, label='Generator Bus', linestyle='None'),
        Line2D([0], [0], marker='o', color='w', markerfacecolor=DARK['load'],
               markersize=7, label='Load Bus', linestyle='None'),
    ]
    ax.legend(handles=legend_els, loc='lower left',
              facecolor=DARK['panel'], edgecolor=DARK['edge'],
              labelcolor=DARK['text'], fontsize=9)

    # System strength annotation
    sys_str = H_B_all.sum()
    ax.set_title(
        f'{name} Network Topology – Bus-Level Inertia Distribution\n'
        f'System Strength = {sys_str:.1f} s   |   '
        f'Max H_B = {h_max:.1f} s (Bus {np.argmax(h_vals)+1})   |   '
        f'Min H_B = {h_min:.1f} s (Bus {np.argmin(h_vals)+1})',
        color=DARK['text'], fontsize=11, pad=10
    )

    fname = f"topology_{name.lower().replace(' ', '').replace('-', '')}.png"
    return _save(fig, fname)


# ── 2. Bus inertia bar chart ──────────────────────────────────────────────────

def plot_bus_inertia_bar(system_data, H_B_all, name):
    """Bar chart of bus-level inertia for all buses."""
    n = system_data.N_BUS
    gen_set = set(system_data.GEN_BUSES)
    buses = np.arange(1, n + 1)

    colors = [DARK['gen'] if b in gen_set else DARK['load'] for b in buses]

    fig, axes = plt.subplots(2, 1, figsize=(max(12, n * 0.22), 9),
                              gridspec_kw={'height_ratios': [3, 1]})
    fig.patch.set_facecolor(DARK['bg'])

    ax = axes[0]
    ax.set_facecolor(DARK['bg'])
    bars = ax.bar(buses, H_B_all, color=colors, edgecolor='none', width=0.75)

    # Highlight top-3 and bottom-3
    sorted_idx = np.argsort(H_B_all)
    for i in sorted_idx[-3:]:
        bars[i].set_edgecolor(DARK['accent'])
        bars[i].set_linewidth(1.5)
    for i in sorted_idx[:3]:
        bars[i].set_edgecolor(DARK['orange'])
        bars[i].set_linewidth(1.5)

    ax.set_xlabel('Bus Number', color=DARK['text'], fontsize=10)
    ax.set_ylabel('Bus Inertia H_B [s]', color=DARK['text'], fontsize=10)
    ax.set_title(f'{name} – Bus-Level Inertia Distribution', color=DARK['text'], fontsize=12)
    ax.tick_params(colors=DARK['text'])
    ax.set_facecolor(DARK['bg'])
    for spine in ax.spines.values():
        spine.set_edgecolor(DARK['panel'])

    if n <= 39:
        ax.set_xticks(buses)

    # System strength line
    sys_str = H_B_all.sum() / n
    ax.axhline(sys_str, color=DARK['accent'], lw=1.2, ls='--',
               label=f'Avg H_B = {sys_str:.1f} s')
    ax.legend(facecolor=DARK['panel'], edgecolor=DARK['edge'],
              labelcolor=DARK['text'], fontsize=9)

    # Cumulative distribution in lower subplot
    ax2 = axes[1]
    ax2.set_facecolor(DARK['bg'])
    sorted_H = np.sort(H_B_all)
    cdf = np.arange(1, n + 1) / n
    ax2.plot(sorted_H, cdf * 100, color=DARK['cyan'], lw=2)
    ax2.fill_between(sorted_H, cdf * 100, alpha=0.15, color=DARK['cyan'])
    ax2.set_xlabel('H_B [s]', color=DARK['text'], fontsize=9)
    ax2.set_ylabel('CDF [%]', color=DARK['text'], fontsize=9)
    ax2.tick_params(colors=DARK['text'])
    for spine in ax2.spines.values():
        spine.set_edgecolor(DARK['panel'])

    plt.tight_layout()
    fname = f"inertia_bar_{name.lower().replace(' ', '').replace('-', '')}.png"
    return _save(fig, fname)


# ── 3. Frequency response comparison ─────────────────────────────────────────

def plot_frequency_response(t, freq_all, system_data, H_B_all, name,
                             highlight_buses=None, f0=60.0):
    """Plot frequency response curves for selected buses.

    Parameters
    ----------
    t            : ndarray (n_pts,)    time [s]
    freq_all     : ndarray (n_bus, n_pts)  frequency [Hz]
    highlight_buses : list[int] or None   1-indexed buses to highlight
    """
    n = system_data.N_BUS
    gen_set = set(system_data.GEN_BUSES)

    if highlight_buses is None:
        # Pick buses with max, min, and median H_B among load buses
        load_idx = [i for i in range(n) if (i + 1) not in gen_set]
        H_load = H_B_all[load_idx]
        i_max = load_idx[np.argmax(H_load)]
        i_min = load_idx[np.argmin(H_load)]
        i_med = load_idx[np.argsort(H_load)[len(H_load) // 2]]
        highlight_buses = [i_max + 1, i_med + 1, i_min + 1]

    fig, axes = plt.subplots(2, 1, figsize=(12, 9),
                              gridspec_kw={'height_ratios': [3, 1]})
    fig.patch.set_facecolor(DARK['bg'])

    ax = axes[0]
    ax.set_facecolor(DARK['bg'])

    palette = [DARK['gen'], DARK['cyan'], DARK['orange'],
               DARK['green'], DARK['accent'], '#cc44cc']

    # Plot all buses lightly
    for i in range(n):
        if (i + 1) not in highlight_buses:
            ax.plot(t, freq_all[i], color='#334466', lw=0.4, alpha=0.35)

    # Highlight selected buses
    for k, bus in enumerate(highlight_buses):
        h_b = H_B_all[bus - 1]
        nadir = freq_all[bus - 1].min()
        ax.plot(t, freq_all[bus - 1], color=palette[k % len(palette)],
                lw=2.2, label=f'Bus {bus}  (H_B={h_b:.1f}s, nadir={nadir:.3f}Hz)')

    ax.axhline(f0, color='white', lw=0.8, ls='--', alpha=0.5)
    ax.axhline(59.5, color=DARK['orange'], lw=1.0, ls=':', alpha=0.8,
               label='UFLS threshold (59.5 Hz)')

    ax.set_xlabel('Time [s]', color=DARK['text'], fontsize=10)
    ax.set_ylabel('Frequency [Hz]', color=DARK['text'], fontsize=10)
    ax.set_title(f'{name} – Frequency Response After Generation Loss\n'
                 f'delta_P = 0.1 pu  |  Bold = highlighted buses',
                 color=DARK['text'], fontsize=11)
    ax.legend(facecolor=DARK['panel'], edgecolor=DARK['edge'],
              labelcolor=DARK['text'], fontsize=9, loc='lower right')
    ax.tick_params(colors=DARK['text'])
    for spine in ax.spines.values():
        spine.set_edgecolor(DARK['panel'])

    # ROCOF subplot
    ax2 = axes[1]
    ax2.set_facecolor(DARK['bg'])
    for k, bus in enumerate(highlight_buses):
        df_dt = np.gradient(freq_all[bus - 1], t)
        ax2.plot(t[:100], df_dt[:100], color=palette[k % len(palette)],
                 lw=1.8, label=f'Bus {bus}')
    ax2.axhline(0, color='white', lw=0.5, ls='--', alpha=0.4)
    ax2.set_xlabel('Time [s]', color=DARK['text'], fontsize=9)
    ax2.set_ylabel('ROCOF [Hz/s]', color=DARK['text'], fontsize=9)
    ax2.tick_params(colors=DARK['text'])
    for spine in ax2.spines.values():
        spine.set_edgecolor(DARK['panel'])

    plt.tight_layout()
    fname = f"freq_response_{name.lower().replace(' ', '').replace('-', '')}.png"
    return _save(fig, fname)


# ── 4. Weighting matrix heatmap ───────────────────────────────────────────────

def plot_weighting_matrix(W_c, system_data, name):
    """Heatmap of the inertia weighting matrix W_c (nb x ng)."""
    gen_buses  = system_data.GEN_BUSES
    n_bus      = system_data.N_BUS
    gen_set    = set(gen_buses)
    load_buses = [b for b in range(1, n_bus + 1) if b not in gen_set]

    # For large systems, subsample load buses
    max_display = 40
    if len(load_buses) > max_display:
        step = len(load_buses) // max_display
        load_buses_disp = load_buses[::step][:max_display]
        row_idx = [load_buses.index(b) for b in load_buses_disp]
        W_disp = W_c[row_idx, :]
    else:
        load_buses_disp = load_buses
        W_disp = W_c

    fig, ax = plt.subplots(figsize=(max(10, len(gen_buses) * 0.7),
                                     max(8, len(load_buses_disp) * 0.22)))
    fig.patch.set_facecolor(DARK['bg'])
    ax.set_facecolor(DARK['bg'])

    im = ax.imshow(W_disp, aspect='auto', cmap='YlOrRd',
                   interpolation='nearest', vmin=0)
    cbar = fig.colorbar(im, ax=ax, fraction=0.03, pad=0.02)
    cbar.set_label('Weight W_c', color=DARK['text'], fontsize=9)
    cbar.ax.yaxis.set_tick_params(color=DARK['text'])
    plt.setp(cbar.ax.yaxis.get_ticklabels(), color=DARK['text'])

    ax.set_xticks(range(len(gen_buses)))
    ax.set_xticklabels([str(b) for b in gen_buses], color=DARK['text'], fontsize=8)
    ax.set_yticks(range(len(load_buses_disp)))
    ax.set_yticklabels([str(b) for b in load_buses_disp], color=DARK['text'], fontsize=7)
    ax.set_xlabel('Generator Bus', color=DARK['text'], fontsize=10)
    ax.set_ylabel('Load Bus', color=DARK['text'], fontsize=10)
    ax.set_title(f'{name} – Inertia Weighting Matrix W_c\n'
                 f'(electrical coupling: load bus vs. generator bus)',
                 color=DARK['text'], fontsize=11)

    plt.tight_layout()
    fname = f"weighting_matrix_{name.lower().replace(' ', '').replace('-', '')}.png"
    return _save(fig, fname)


# ── 5. RES penetration impact ─────────────────────────────────────────────────

def plot_res_impact(res_results, name):
    """Bar/line chart showing system strength vs. RES penetration."""
    pcts    = [r['res_pct'] for r in res_results]
    sys_str = [r['sys_str'] for r in res_results]
    n_on    = [r['n_online'] for r in res_results]

    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    fig.patch.set_facecolor(DARK['bg'])

    # System strength vs RES
    ax = axes[0]
    ax.set_facecolor(DARK['bg'])
    ax.bar(pcts, sys_str, color=DARK['load'], edgecolor='none',
           width=8, alpha=0.85)
    ax.plot(pcts, sys_str, 'o-', color=DARK['accent'], lw=2, ms=6)
    ax.set_xlabel('RES Penetration [%]', color=DARK['text'], fontsize=10)
    ax.set_ylabel('System Strength [s]', color=DARK['text'], fontsize=10)
    ax.set_title(f'{name}\nSystem Strength vs. RES Penetration',
                 color=DARK['text'], fontsize=11)
    ax.tick_params(colors=DARK['text'])
    for spine in ax.spines.values():
        spine.set_edgecolor(DARK['panel'])

    # Online generators vs RES
    ax2 = axes[1]
    ax2.set_facecolor(DARK['bg'])
    ax2.step(pcts, n_on, color=DARK['gen'], lw=2.5, where='post')
    ax2.fill_between(pcts, n_on, step='post', alpha=0.2, color=DARK['gen'])
    ax2.set_xlabel('RES Penetration [%]', color=DARK['text'], fontsize=10)
    ax2.set_ylabel('Online Synchronous Generators', color=DARK['text'], fontsize=10)
    ax2.set_title(f'{name}\nOnline Generators vs. RES Level',
                  color=DARK['text'], fontsize=11)
    ax2.tick_params(colors=DARK['text'])
    for spine in ax2.spines.values():
        spine.set_edgecolor(DARK['panel'])

    plt.tight_layout()
    fname = f"res_impact_{name.lower().replace(' ', '').replace('-', '')}.png"
    return _save(fig, fname)


# ── 6. System comparison summary ──────────────────────────────────────────────

def plot_system_comparison(all_results):
    """Bar chart comparing system strength across all IEEE test systems."""
    systems  = [r['name'] for r in all_results]
    sys_strs = [r['sys_str'] for r in all_results]
    n_gens   = [r['n_gen'] for r in all_results]
    n_buses  = [r['n_bus'] for r in all_results]

    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    fig.patch.set_facecolor(DARK['bg'])

    palette = [DARK['gen'], DARK['load'], DARK['cyan'], DARK['green']]

    ax = axes[0]
    ax.set_facecolor(DARK['bg'])
    x = np.arange(len(systems))
    bars = ax.bar(x, sys_strs, color=palette[:len(systems)], edgecolor='none', width=0.55)
    for bar, val in zip(bars, sys_strs):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 50,
                f'{val:.0f}s', ha='center', color=DARK['text'], fontsize=10)
    ax.set_xticks(x)
    ax.set_xticklabels(systems, color=DARK['text'], fontsize=10)
    ax.set_ylabel('Total System Strength [s]', color=DARK['text'], fontsize=10)
    ax.set_title('System Inertia Comparison\nAll IEEE Test Systems',
                 color=DARK['text'], fontsize=11)
    ax.tick_params(colors=DARK['text'])
    for spine in ax.spines.values():
        spine.set_edgecolor(DARK['panel'])

    ax2 = axes[1]
    ax2.set_facecolor(DARK['bg'])
    avg_h = [s / b for s, b in zip(sys_strs, n_buses)]
    bars2 = ax2.bar(x, avg_h, color=palette[:len(systems)], edgecolor='none', width=0.55)
    for bar, val in zip(bars2, avg_h):
        ax2.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                 f'{val:.1f}s', ha='center', color=DARK['text'], fontsize=10)
    ax2.set_xticks(x)
    ax2.set_xticklabels(systems, color=DARK['text'], fontsize=10)
    ax2.set_ylabel('Avg Bus Inertia H_B [s/bus]', color=DARK['text'], fontsize=10)
    ax2.set_title('Average Bus-Level Inertia\nAll IEEE Test Systems',
                  color=DARK['text'], fontsize=11)
    ax2.tick_params(colors=DARK['text'])
    for spine in ax2.spines.values():
        spine.set_edgecolor(DARK['panel'])

    plt.tight_layout()
    return _save(fig, 'system_comparison_all.png')
