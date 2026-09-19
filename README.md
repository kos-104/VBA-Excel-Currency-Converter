# Excel Currency Converter (VBA)

**A responsive, object-oriented Excel tool for retrieving and analyzing real-time currency data.**

![VBA](https://img.shields.io/badge/Language-VBA-blue.svg)
![Excel](https://img.shields.io/badge/Platform-Excel-217346.svg)
![OOP](https://img.shields.io/badge/Design-OOP-orange.svg)
![API](https://img.shields.io/badge/Feature-Live%20API%20Data-brightgreen.svg)

## Description

This Excel Currency Converter (developed with VBA) returns currency information online.

It features complex event handling and event-trapping to ensure smooth communication between program components. Class modules help keep the code organized and encapsulated, within an object-oriented approach. Robust error-handling also captures expected and [most] unexpected slips.

The form seeks to be intuitive, allowing users to launch it quickly and access scrollable currency dropdown-menus with varied sorting options. A collapsible interactive calendar permits users to select dates for analyzing currency trends over the last 30 days. Users get real-time feedback with the form and can cancel a process and/or get partial results where available.

Many API calls are utilized, including some that can position the form on any worksheet cell, and others that enable the user to resize the userform and its controls.

---

## Version History

### 🔹 Version 92 (Original)

- Userform could be **freely resized and repositioned** on any worksheet cell via API calls
- Employed reasonable modularity, but relied more heavily on **global variables**, increasing coupling between components
- Combined larger responsibilities within single procedures, which complicated testing and maintenance
- Used **WithEvents objects** for event handling
- Relied on public variables for data access rather than properties

### 🔹 Version 2.3 (Rebuild)

In this version, we've increased the efficiency of the design:

- To accelerate performance and limit resource-drain, users can **no longer resize or reposition** the form on screen
- **Dictionaries** now manage code information, allowing for faster lookups and easier array output while minimizing range calls to the worksheet
- Greater use of **encapsulation, refactoring, and modularity**, with private variables and focused subroutines protecting internal states from outside dependencies
- Procedures are more focused and single-purpose, simplifying testing and reuse
- Event-handling shifted from **WithEvents objects to Custom event objects**, decoupling event sources from their handlers for greater flexibility
- Expanded use of **OOP principles** — abstraction and classification via class modules and custom types encapsulate data and behavior
- **Properties** replace public variables, supporting data hiding, validation, data integrity, and control
- Implements numerous code-optimization suggestions provided by the **Rubberduck VBA add-in**

---

## Comparison Summary

| Aspect               | Version 92 (Original)         | Version 2.3 (Rebuild)                  |
|----------------------|--------------------------------|------------------------------------------|
| Form Resize/Reposition | ✅ Enabled via API calls      | ❌ Disabled for performance             |
| Data Management      | Range-based lookups            | Dictionary-based lookups                |
| Variable Scope        | Global variables (higher coupling) | Private variables + properties (lower coupling) |
| Procedure Design      | Larger, multi-responsibility procedures | Smaller, focused procedures            |
| Event Handling        | WithEvents objects             | Custom event objects                    |
| OOP Usage             | Reasonable modularity           | Stronger abstraction & classification   |
| Code Optimization     | —                               | Rubberduck VBA-guided improvements      |
