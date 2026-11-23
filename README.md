# STAAD Tank Wall Explorer

A desktop tool (Python + OpenSTAAD) for post-processing **circular water tank** models in STAAD.Pro.
It reads the plate results from an analysed model and converts the **local-axis** plate moments and
in-plane forces into **design-ready global quantities**:

- **Vertical bending moment** (kN·m/m) up the tank wall
- **Hoop (circumferential) force** (kN/m)
- **Full cross-section diagrams**: BMD and hoop force drawn around the whole tank section through its centre (domes, wall, cone)

![Cross section example](images/cross_section_example.png)

## Why

A meshed tank uses many triangular plates, and each one has its own local axis orientation. STAAD reports
`Mx`, `My`, `Mxy`, `Sx`, `Sy` and `Sxy` in those local axes, so they cannot be read directly as vertical or hoop values.
This tool rotates them to the required direction with the standard transformation (Mohr's circle):

```
M(θ) = (Mx + My)/2 + (Mx − My)/2 · cos2θ + Mxy · sin2θ
N_hoop = σ(α) × plate thickness
```

## Features

- Pick any **load case or load combination** (LC 1 to LC 20 in the list, or type any number such as 201)
- Select a vertical strip by **plane angle ± tolerance** and an **elevation range**
- **Bin** plates by elevation (Bin Size) to get one diagram: average, envelope (min/max) or raw plates
- **Tank Cross Section** tab: BMD and hoop force along the full meridian on both sides of the cut, with optional values
- 3D view of the whole tank with the selected plates highlighted (PyVista)
- Selected-plate table and a step-by-step **calculation report**
- **Export to Excel** (data, bins, cross-section stations, charts) and save diagrams as PNG

## Requirements

- Windows with **STAAD.Pro** (CONNECT Edition / 2023+) that includes OpenSTAAD Python (`openstaadpy`)
- Python 3.10+

```bash
pip install -r requirements.txt
```

## Usage

1. Open your tank model in STAAD.Pro and **run the analysis**.
2. Run the tool:
   ```bash
   python STAAD_Tank_Wall_Explorer.py
   ```
3. Set the result type, elevation range, load case number, plane angle, tolerance and bin size.
4. Click **EXTRACT AND DISPLAY FULL TANK**.
5. Review the tabs (Interactive Diagram, Tank Cross Section, 3D view, table, report), then **Save Diagram** or **Export to Excel**.

## Sign conventions

- Positive plate moment = tension on the plate's local +Z face.
- In the cross-section view, the moment is drawn on the **tension face**, and hoop tension is drawn outside the outline.
- The tool assumes the tank axis is vertical (global Y) through X = 0, Z = 0.

## Notes and limitations

- Ring beams are frame members, so their hoop tension is not included in the plate results.
- With an unstructured mesh, bins near the dome crown may contain only 1–3 plates, which can give small differences between the two sides of the section.
- Check the results against hand calculations before using them for design.

## Author

**Biswatosh**, Bentley Systems

YouTube channel: **TORQ | Simplifying Engineering**

## License

MIT, see [LICENSE](LICENSE).
