# InDesign Figure Automation

An Adobe InDesign ExtendScript that automates the placement, positioning, anchoring, and captioning of figures in editorial documents.

The script is designed to reduce repetitive layout work by automatically matching figures to their references in the document text and placing each figure together with its corresponding caption.

## Overview

Manually inserting and positioning figures in long documents can be time-consuming and error-prone, particularly when a document contains a large number of figures and references.

This project automates that workflow.

The script:

1. Reads a directory containing the figures and their corresponding captions.
2. Searches the InDesign document for references to each figure.
3. Inserts the corresponding figure at the appropriate location.
4. Anchors the figure to the relevant text.
5. Creates and positions the figure caption.
6. Applies the configured paragraph style to the caption.
7. Groups the figure and caption so they can be treated as a single layout element.
8. Repeats the process for the selected figures.

The result is a more consistent and efficient figure-placement workflow for Adobe InDesign.

## Features

* **Automatic figure matching**
  Finds figure references in the document and associates them with the corresponding image.

* **Automatic placement**
  Inserts figures at the location of their references.

* **Anchored figures**
  Figures are anchored to the relevant text so that they remain associated with the corresponding content.

* **Automatic captions**
  Captions are generated and placed underneath their associated figures.

* **Paragraph style support**
  Captions can be formatted using a predefined InDesign paragraph style.

* **Figure + caption grouping**
  The figure and its caption are grouped together to simplify subsequent layout adjustments.

* **Batch processing**
  The workflow can be repeated across multiple figures rather than requiring manual placement of each one.


## Requirements

* Adobe InDesign
* Adobe ExtendScript-compatible scripting environment
* An InDesign document containing figure references
* A directory containing the corresponding figures

The scripts use Adobe InDesign's scripting API to interact directly with the document and its layout objects.

## Motivation

The project was created to automate a common editorial production task: maintaining the relationship between figures, figure references, and captions in an InDesign document.

Instead of manually:

* searching for every figure reference,
* locating the corresponding image,
* placing the image,
* positioning it,
* creating its caption,
* formatting the caption, and
* keeping the two elements together,

the process can be handled programmatically.

## Technology

* **Adobe InDesign**
* **ExtendScript**
* **JavaScript**
* **InDesign DOM / scripting API**

The project demonstrates how scripting can be used to extend InDesign's native capabilities and automate complex editorial workflows.
