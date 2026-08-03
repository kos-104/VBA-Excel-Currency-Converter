This Excel Currency Converter (developed with VBA) returns currency information online.

It features complex event handling and event-trapping to ensure smooth communication between program components. Class modules help keep the code organized and encapsulated, within an object-oriented approach. Robust error-handling also captures expected and [most] unexpected slips.

The form seeks to be intuitive, allowing users to launch it quickly and access scrollable currency dropdown-menus with varied sorting options. A collapsible interactive calendar permits users to select dates for analyzing currency trends over the last 30 days. Users get real-time feedback with the form and can cancel a process and/or get partial results where available.

Many API calls are utilized, including some that can position the form on any worksheet cell, and others that enable the user to resize the userform and its controls.
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
In the 2nd version (Currency Converter Rebuild 2.3), we've increased the efficiency of the design.

To accelerate performance and limit resource-drain, users can no longer resize nor reposition the form on screen. Dictionaries now manage code information, allowing for faster lookups and easier array output while minimizing range calls to the worksheet.

This version provides greater demonstrations of encapsulation, refactoring, and modularity. It presents better usage of private variables and specific subroutines, enhancing encapsulation by protecting internal states against outside dependencies. This contrasts with Version 92 - which employed reasonable modularity - but relied more on global variables that led to increased coupling. 

The latest version’s procedures are more focused, simplifying testing and reuse; versus the original which combined larger responsibilities within single procedures, thereby complicating maintenance.

The event-handling protocol has shifted from WithEvents objects to Custom event objects, increasing flexibility by decoupling event sources from their handlers. 

This version makes greater use of certain OOP (object-oriented-principles) techniques through abstraction and classification, with class modules and custom types encapsulating data and behavior. It also deploys properties instead of public variables to support data hiding, validation, data integrity, and control.

The current version also implements many code-optimizing suggestions provided by the Rubberduck VBA add-in.
