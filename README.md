# TROA Scheduling PowerShell Scripts

# Overview

This repository contains PowerShell scripts designed for interacting with the TROA Scheduling SharePoint lists. The scripts allow for test development on clones of the scheduling lists on the PWRE SharePoint site, as well as official development on the scheduling lists hosted by the water master (USWM). Any user with edit access on the USWM site and PWRE site will be able to use these scripts with their existing MS logins.

# Installation/Requirements

- Install [SharepointPnPPowerShellOnline](https://pnp.github.io/powershell/):
  - Run PowerShell as administrator

    ```Install-Module SharePointPnPPowerShellOnline```

- Set execution policy to allow scripts to run

    ```Set-ExecutionPolicy RemoteSigned```

- Create the local authorization file
  - Web logins are used for the PWRE site
  - For the USWM site, a common service account is used. Create an authorization file for this account called "USWM.auth" (will be ignored by GitHub) and save it in the "Functions" directory. The screenshot below shows the required file format

![](img/AuthFile.png?raw=true)

_Figure 1: USWM authorization file format_

# Instructions

## Main GUI

For any scheduling related development, begin by running the top-level script, "TROAScheduling\_Main.ps1". This launches a GUI with three options:

- Schedule Integration
- Schedule List Review
- Copy Schedule List(s) to Test Site

Each option represents a category of scheduling list development and will launch a subsequent GUI, which are each explained in detail in their own following subsections.

![](img/MainGUI.png?raw=true)

_Figure 2: Main GUI_

## Schedule Integration

The Schedule Integration GUI, shown below, is used to control integration at monthly TROA meetings. It can also be used to test integration on PWRE test SharePoint lists.

![](img/IntegrationGUI.png?raw=true)

_Figure 3: Schedule Integration GUI_

- Select which site to perform the integration on
- Select the scope of integration
  - All lists (default)
  - Single party's lists
  - Single list
  - Specific item on a single list
- The "Single List" or "Specific item(s)" radio buttons will enable the list name text box to specify a list
- If the "Specific item(s)" radio is selected, "OK" will launch a window to select one or multiple items to integrate (shown below). Otherwise, the integration will occur with no further input

![](img/ItemSelection.png?raw=true)

_Figure 4: Selection window for specific item integration_

## Schedule List Review

The Schedule List Review GUI, shown below, is used to perform various checks on the scheduling lists for QA/QC or other information.

![](img/ListReviewGUI.png?raw=true)

_Figure 5: Schedule List Review GUI_

- Select which site to perform the integration on
- Press "Load Data" to load all current lists, list items, and metadata from SharePoint into memory
  - This takes a couple minutes, but once complete multiple different reviews of different types can be done instantly, by using the data in memory rather than having to reload data from SharePoint
- Select the review type
  - Count Recently Modified Items
    - Returns a window showing the number of recently modified items for each list
    - Requires an input of a Recently Modified Date threshold
  - Count Integrated and Draft Items
    - Returns a window showing the number of integrated and draft items for each list
  - Check for Unsubmitted Draft Items
    - Returns a window showing the number (if any) of unsubmitted draft items for each list
  - Edit Modified Date of Specific Item(s)
    - Launches a window allowing the user to select one or multiple list items to manually adjust the "Recently Modified" metadata field
    - Requires an input for the desired "Recently Modified" value (text input box header will change if this review type is selected)
    - Requires an input for the list containing items to be modified (text input box will be activated if this review type is selected)

A sample output window is shown below.

![](img/OutputWindow.png?raw=true)

_Figure 6: Sample output window, using "Count Integrated and Draft Items" option_

## Copy Lists to Test Site

The Copy Lists GUI, shown below, is used to copy one or all scheduling lists with items and metadata, from the official USWM site to PWRE's test site. This can be useful for testing new features of the PowerShell scripts, or for testing any future projects involving scheduling automation.

![](img/CopyListGUI.png?raw=true)

_Figure 7: Copy List GUI_

- Select copy scope
  - All Lists
  - Single List
- If Single List scope is selected, the text input box will be activated for the user to specify a single list to be copied