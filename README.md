# Efficient Portfolio

A simple and easy-to-use Python application for analyzing and optimizing stock portfolios. This tool helps users fetch historical stock data from Yahoo Finance, calculate various financial metrics, and visualize the efficient frontier of a portfolio. It includes a user-friendly GUI for beginners.

## How-To-Setup

1. **Install Python:**
   - For macOS/Linux, download Python from [python.org](https://www.python.org/downloads/).
   - For Windows, it's best to download the latest version from the Microsoft Store.

2. **Install Required Libraries:**
   - You can install the necessary libraries listed in the `requirements.txt` file. To do this, place the `requirements.txt` file in the same folder as the `.py` file and run:
     ```bash
     pip install -r requirements.txt
     ```
   - Alternatively, you can install the libraries as needed based on any error messages that arise.

3. **Download the Code:**
   - Download the `main.py` file from the repository. Move this file to a folder on your computer. You can rename it if you wish, but keep it simple.

4. **Open a Terminal:**
   - **Linux/macOS:** Use `cd` to navigate to the directory. Example:
     ```bash
     cd /path/to/your/directory
     ```
   - **Windows:** Open Command Prompt or PowerShell, then use `cd` to navigate to the directory. Example:
     ```cmd
     cd C:\path\to\your\directory
     ```

5. **Run the Application:**
   - Execute the following command:
     ```bash
     python whatyounamedit.py
     ```

6. **Handle Missing Libraries:**
   - If you encounter errors about missing modules, install them as instructed by the error messages.

7. **Using the GUI:**
   - After running the application, a GUI will appear. Follow the instructions in the "How-To-Use" section to interact with it.

## How-To-Use

1. **Run the Application:**
   - Refer to step #5 in the "How-To-Setup" section.

2. **Enter Stock Symbols:**
   - Input stock symbols separated by commas or use the search function to find stocks by company name.

3. **Select Date Range:**
   - Choose the start and end dates for the historical data.

4. **Set Interval and Risk-Free Rate:**
   - Select the data interval (daily, weekly, monthly).
   - Input the annual risk-free rate as a percentage (e.g., 2 for 2%).

5. **Fetch Data:**
   - Click "Fetch Data" to retrieve and analyze the data.

6. **Save Results:**
   - Name the Excel file and click "OK" to save the results.

## Results

- Navigate to the directory where you ran the code. You will find an Excel file named as you specified.
- The stock data includes open, high, low, close, adjusted closing, volume, efficient frontier, optimal portfolio, correlation matrix, and covariance matrix.

## Thanks

If you have suggestions or improvements, please let me know, and I'll consider them for future updates.
