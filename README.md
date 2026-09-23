# Excel VBA Shape Animator

**A checkbox-driven 3D animation demo for Excel — proving that Microsoft Excel is capable of more than just spreadsheets.**

![VBA](https://img.shields.io/badge/Language-VBA-blue.svg)
![Excel](https://img.shields.io/badge/Platform-Excel-217346.svg)
![Animation](https://img.shields.io/badge/Feature-3D%20Rotation-brightgreen.svg)
![Dependencies](https://img.shields.io/badge/Dependencies-Scripting%20Runtime-yellow.svg)

## Description

Most people think of Excel as a grid for numbers. This project was built to challenge that assumption — showing that, backed entirely by VBA, Excel can render live 3D graphics, colorful shading, and animation using only its native engine.

Users generate shapes from dropdown selections, assign them a fill pattern, and set them spinning via multiple rotation modes — controlled by simple checkboxes. Shape creation resides within a dedicated class module that listens for worksheet events, responding live after the user types into a cell. A separate module handles the animation itself, adding depth and shadows to make shapes appear three-dimensional, shifting their colors frame by frame, with a graphical stop button to pause any running animations.

This was built as a passion project, meant to be a fun, visual reminder that Excel can be appreciably playful when you venture beyond its default look and feel.

## Key Features

**Shape Creation**
- Pick a shape from a dropdown list — circles, diamonds, arrows, and more — no drawing required
- Every shape gets a randomized, multi-color blend fill, so no two shapes look identical
- Choose how big the shape should be: a percentage you type in, from 1 to 100, with out-of-range or negative values correcting automatically

**Animation**
- Seven spinning modes, toggled with checkboxes — spin on one axis, two at once, or all three together, with speed and direction varying as rotations accumulate
- Shading and shadow effects that make the spin look genuinely 3D instead of flat
- Colors shift a little on every frame, so the animation never looks static or repetitive
- Start and stop the animation with a click, with no freezing or waiting
- Toggle between "video" mode and the default Excel view with a single button

**Under the Hood**
- Shape creation is built as a primitive object-oriented model — a self-contained class module that responds directly to worksheet events
- Shape creation and animation are kept in separate pieces of code, keeping event-driven logic and animation logic cleanly apart

## Requirements
- Excel with macros enabled (developed and tested on Excel 365)
- The following VBA Reference must be enabled (**Tools → References** in the VBA Editor):
  - Microsoft Scripting Runtime

## How to Use
1. Open the provided `.xlsm` file
2. Follow the on-screen instructions provided directly on the Excel sheet

---

*Built to reveal that Microsoft Excel is full of surprises.*
