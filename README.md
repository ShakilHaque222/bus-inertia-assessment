# Bus-Level Inertia Assessment for Power Systems

Implementation of the bus-level inertia assessment algorithm from:

> **Ghosh, S., Isbeih, Y. J., & El Moursi, M. S. (2023).**  
> *Bus-Level Inertia Assessment for Converter-Dominated Power Systems.*  
> IEEE Transactions on Power Delivery, Vol. 38, No. 4, pp. 2635–2648.

Covers IEEE 14-bus, 39-bus (New England), 68-bus (NETS-NYPS), and 118-bus test systems.

---

## Algorithm Overview

### Core Equations

**Weighting Matrix** (Eq. 15):

```
W_c(t) = { -Y_bb⁻¹(t) × Y_bg(t) ./ v_b(t) } .* e_g(t)
```

**Bus-Level Inertia** (Eq. 14):

```
H_B(t) = W_c(t) × H_G
```

**System Strength** (Eq. 16):

```
Sys_str(t) = Σ H_B(t)
```

Where:
- `Y_bb` — admittance submatrix among non-generator (load) buses
- `Y_bg` — admittance submatrix from load buses to generator buses
- `v_b`  — voltage magnitudes at load buses [pu]
- `e_g`  — binary generator participation vector (1=online, 0=tripped)
- `H_G`  — generator inertia constants [seconds]

### Fault Condition

During a fault, bus voltages are assumed flat (`|v_b| = 1 pu`):

```
W_c = -Y_bb⁻¹ × Y_bg
```

---

## Project Structure

```
bus_inertia_project/
├── data/
│   ├── ieee14.py      # IEEE 14-bus (5 gens)
│   ├── ieee39.py      # IEEE 39-bus New England (10 gens)
│   ├── ieee68.py      # IEEE 68-bus NETS-NYPS (16 gens)
│   └── ieee118.py     # IEEE 118-bus (54 gens)
├── core/
│   ├── ybus.py        # Y-bus construction and partitioning
│   ├── bus_inertia.py # Main algorithm (Eqs. 14-16)
│   └── freq_sim.py    # Swing equation frequency simulation
├── viz/
│   └── plots.py       # All figure generation
├── results/           # Output PNG figures
├── main.py            # Main runner
└── requirements.txt
```

---

## Installation

```bash
pip install -r requirements.txt
```

**Requirements:** numpy, scipy, matplotlib, networkx

---

## Usage

```bash
# Run all 4 IEEE test systems
python main.py

# Run a specific system
python main.py --system 39

# Skip frequency simulation (faster)
python main.py --no-freq

# Run N-1 contingency analysis
python main.py --system 14 --contingency

# All options
python main.py --help
```

---

## Results

### Generated Figures (per system)

| Figure | Description |
|--------|-------------|
| `topology_ieee*bus.png` | Network topology with bus inertia colour map |
| `inertia_bar_ieee*bus.png` | Bus-level inertia bar chart + CDF |
| `freq_response_ieee*bus.png` | Frequency response after generation loss |
| `weighting_matrix_ieee*bus.png` | W_c matrix heatmap (electrical coupling) |
| `res_impact_ieee*bus.png` | System strength vs. RES penetration |
| `system_comparison_all.png` | Cross-system inertia comparison |

### Sample Numerical Results

| System | Buses | Gens | System Strength | Avg H_B |
|--------|-------|------|----------------|---------|
| IEEE 14-Bus | 14 | 5 | 75.3 s | 5.4 s |
| IEEE 39-Bus | 39 | 10 | 2860 s | 73.3 s |
| IEEE 68-Bus | 68 | 16 | 7201 s | 105.9 s |
| IEEE 118-Bus | 118 | 54 | 653 s | 5.5 s |

### Key Observations

1. **Electrical proximity matters** – Load buses electrically closest to large generators have the highest H_B values.
2. **RES penetration reduces H_B** – Replacing synchronous generators with converter-interfaced RES lowers system strength monotonically.
3. **ROCOF correlates inversely with H_B** – Low-inertia buses experience faster frequency drops after generation loss events.
4. **IEEE 39-bus Bus 1** has the highest load-bus inertia (directly coupled to the large equivalent generator at Bus 39, H_G=500s).

---

## Physical Interpretation

The bus-level inertia H_B(bus i) represents the **effective inertia seen at bus i**, computed as a weighted sum of all generator inertia constants. The weights W_c are derived from the inverse admittance matrix and capture the **electrical distance** between each load bus and each generator.

High H_B → bus is electrically close to large generators → slow frequency response  
Low  H_B → bus is electrically distant from generators → fast frequency drop (high ROCOF)

---

## Reference

```bibtex
@article{ghosh2023bus,
  author  = {Ghosh, Soumya and Isbeih, Younes Jundi and El Moursi, Mohamed Shawky},
  title   = {Bus-Level Inertia Assessment for Converter-Dominated Power Systems},
  journal = {IEEE Transactions on Power Delivery},
  volume  = {38},
  number  = {4},
  pages   = {2635--2648},
  year    = {2023},
  doi     = {10.1109/TPWRD.2023.3237879}
}
```

---

*Author: Shakil Haque | Email: shakil.haqueee@gmail.com*
