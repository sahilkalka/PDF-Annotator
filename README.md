# PDF Annotation Studio

A lightweight, native desktop tool designed for mapping, verifying, and annotating Test Point and Connector data onto Assembly PDFs. Built with Python, Flask, and PyWebView.

## 🚀 Features
* **Smart Parsing:** Extracts net and pin data directly from `VT Pins` Excel tables.
* **Dual Assembly Support:** Handles `Master` and `Slave` PDF overlays in a single session.
* **Visual Styling Engine:** Customize ring colors, label sizes, text spacing, and background collision layers.
* **Live Preview:** Instant, interactive visual feedback natively rendered in a desktop window.

## 📋 Prerequisites
* Windows 10 or 11
* [Python 3.9+](https://www.python.org/downloads/) (Make sure to check "Add Python to PATH" during installation)

## 🛠️ Installation & Setup

We have provided a one-click setup script that handles all environment configurations automatically.

1. Download or extract the project folder to your desired location.
2. Double-click the `setup.bat` file.
3. The script will:
   * Create an isolated Python virtual environment (`venv`).
   * Install necessary dependencies (`flask`, `PyMuPDF`, `openpyxl`, `pywebview`).
   * Generate an `icon.ico` file if you haven't provided one.
   * Add a **PDF Annotation Studio** shortcut directly to your Windows Start Menu.

*(Optional)* Replace the generated `icon.ico` file in the root folder with your own custom 16x16 or 32x32 `.ico` to personalize your Start Menu and Taskbar presence.

## 🏃‍♂️ Running the Application

After running `setup.bat` for the first time, you can launch the program in two ways:
1. **The easy way:** Open your Windows Start Menu and search for **PDF Annotation Studio**.
2. **The manual way:** Double-click `launch_app.bat` inside the project folder.

## 🖱️ Usage Guide

1. **Upload Files:** Use the `Inputs` panel on the left to drag-and-drop your Master Assembly PDF, Slave Assembly PDF, and VT Pins Excel file.
2. **Select Project:** Once the Excel file is processed, choose your Target Project from the dropdown menu and click **Load**.
3. **Verify Data:** Check the `M-Nets`, `S-Nets`, and `Orphan` tabs on the left to ensure the correct net labels have been identified.
4. **Style & Adjust:** 
   * Use the `Style` and `Advanced` panels to alter colors, font sizes, and ring constraints. 
   * If a label collides with surrounding geometry in the preview, click and **drag the label** directly on the canvas to assign a custom manual offset.
5. **Export:** Click the **Generate Final PDFs** button in the top right header to process and download your annotated documents.