# Odin Delta Trading Tools

A high-speed trading companion for the Odin trading platform. This tool provides real-time premium overlays and is being extended to support lightning-fast order execution via hotkeys and a virtual keyboard.

## 🚀 Key Features

*   **Real-time Premium Overlays:** Displays BSC and NSC premiums directly over your trading screen.
*   **Floating Clock:** A highly visible, always-on-top clock for precise trade timing.
*   **Excel Integration:** Automatically pulls live data from your existing Excel-based delta calculators.
*   **System Tray Control:** Easily toggle overlays and settings from the Windows system tray.

## 🛠 Prerequisites

*   **Python 3.10+**
*   **Microsoft Excel:** Must be running and have your delta calculation sheet open.
*   **Odin Trading Software:** The tool is designed to overlay on top of Odin.

## 📦 Installation

1.  **Clone the directory** or download the source files.
2.  **Install Dependencies:**
    ```bash
    pip install PyQt5 pypiwin32
    ```

## 📖 How to Use

### 1. Launching the Tool
Run `main.py` using Python:
```bash
python main.py
```
A blue "O" icon will appear in your system tray (bottom right corner of Windows).

### 2. Configuring Excel Data
1.  Right-click the "O" icon in the system tray.
2.  Select **Configure BSC/NSC Excel...**.
3.  **Browse** for your Excel file.
4.  Specify the **Sheet Name** and the **Cell Addresses** (e.g., E56) where your premiums are calculated.
5.  Click **Save All Settings**.

### 3. Toggling Overlays
*   **Show Time Overlay:** Displays a red digital clock at the bottom right.
*   **Show BSC/NSC Premium:** Displays the values from your Excel cells in a translucent box.

## ⏳ Coming Soon (Next Update)

*   **High-Speed Hotkeys:** Press `F1`/`F2` to instantly trigger "Buy" or "Sell" in Odin with pre-filled quantities.
*   **Virtual Keyboard:** A floating UI with large buttons for "Square Off", "Buy High", and "Sell Low" to perform actions without moving your mouse to Odin.
*   **Panic Button:** Press `Esc` to instantly stop all automated trading actions.

---
**Disclaimer:** This tool is for informational and automation purposes only. Trading involves significant risk. Ensure you test all automated sequences in a paper-trading environment before live execution.
