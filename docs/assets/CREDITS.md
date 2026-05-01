# Texture & graphic credits

## Background paper textures

**Subtle Patterns / Transparent Textures** — https://www.transparenttextures.com
Author: Mike Hearn (@mikehearn). License: [CC BY 3.0](https://creativecommons.org/licenses/by/3.0/).

| File | Source URL |
|---|---|
| `groovepaper.png`  | https://www.transparenttextures.com/patterns/groovepaper.png |
| `dust.png`         | https://www.transparenttextures.com/patterns/dust.png |
| `cream-paper.png`  | https://www.transparenttextures.com/patterns/cream-paper.png |

## Grunge JPGs (heavy texture overlays) and hand-drawn scribble PNGs

**Resource Boy** — https://resourceboy.com/
License: free for personal and commercial use (per Resource Boy site terms).

Full packs: [scribble textures](https://resourceboy.com/textures/scribble-textures/),
[grunge textures](https://resourceboy.com/textures/grunge-textures/).

| File | Pack source |
|---|---|
| `grunge-heavy.jpg`     | Resource Boy Grunge Textures, image 001 |
| `grunge-speckle.jpg`   | Resource Boy Grunge Textures, image 050 |
| `grunge-light.jpg`     | Resource Boy Grunge Textures, image 100 |
| `grunge-wide.jpg`      | Resource Boy Grunge Textures, image 150 |
| `grunge-mid.jpg`       | Resource Boy Grunge Textures, image 200 |
| `scribble-loops.png`   | Resource Boy Scribble Textures, image 050 |
| `scribble-tangle.png`  | Resource Boy Scribble Textures, image 100 |
| `scribble-zigzag.png`  | Resource Boy Scribble Textures, image 200 |
| `scribble-marker.png`  | Resource Boy Scribble Textures, image 300 |
| `scribble-bold-n.png`  | Resource Boy Scribble Textures, image 450 |
| `arrow-left.png`        | Resource Boy Hand Drawn Arrow Elements, image 044 |
| `arrow-wave-down.png`   | Resource Boy Hand Drawn Arrow Elements, image 227 |
| `arrow-zigzag.png`      | Resource Boy Hand Drawn Arrow Elements, image 332 |
| `leak-warm.jpg`         | Resource Boy Light Leak Overlays, image 033 (warm orange) |
| `leak-cool.jpg`         | Resource Boy Light Leak Overlays, image 005 (green/teal) |

Originals were 4K+ resolution; resized to ~1200–1920px max dimension for the
grunge JPGs and ~900–1400px for the scribble PNGs, then PNG-optimized.
Scribbles are used in CSS as `mask-image` so the shape (real hand-drawn
graphite/marker texture) takes the active palette color. Grunge JPGs are
applied as `mix-blend-mode: screen` overlays so the white speckle/scratch
content paints on top of the underlying surface and the black background
becomes invisible.
