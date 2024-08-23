---
title: "Efficient Portfolio"
author: "Your Name"
date: "`r Sys.Date()`"
output: 
  html_document:
    toc: true
    toc_depth: 2
    theme: united
    highlight: tango
---

# Efficient Portfolio

The **Efficient Portfolio** project is a Python-based application designed to analyze and optimize stock portfolios. This tool helps users fetch historical stock data, calculate various financial metrics, and visualize the efficient frontier of a portfolio.

## Table of Contents

1. [Features](#features)
2. [Installation](#installation)
3. [Usage](#usage)
4. [Code Overview](#code-overview)
5. [Contributing](#contributing)
6. [License](#license)

## Features

- **Provides a GUI for User Input:** The application features a graphical user interface (GUI) created with `tkinter`. Users can input stock symbols, select date ranges, choose data intervals, and manually input an annual risk-free rate.
- Fetches historical stock data from Yahoo Finance.
- Calculates and saves log returns, covariance matrix, and correlation matrix to Excel.
- Generates an efficient frontier and optimal portfolio.

## Installation

To get started with **Efficient Portfolio**, follow these steps:

1. **Clone the Repository:**

    ```bash
    git clone https://github.com/yourusername/efficient-portfolio.git
    cd efficient-portfolio
    ```

2. **Set Up a Virtual Environment:**

    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3. **Install Required Packages:**

    Install the necessary Python packages using `pip`:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. **Run the Application:**

    Start the GUI application by executing:

    ```bash
    python main.py
    ```

2. **Using the GUI:**

    - **Enter Stock Symbols:** Input the stock symbols (e.g., AAPL, MSFT) separated by commas.
    - **Select Date Range:** Choose the start and end dates for the historical data.
    - **Set Interval:** Choose the data interval (daily, weekly, or monthly).
    - **Enter Risk-Free Rate:** Input the annual risk-free rate as a percentage (e.g., 2 for 2%). This rate will be used to calculate the Capital Allocation Line (CAL) and the optimal portfolio.
    - **Fetch Data:** Click the button to start fetching data.
    - **View Results:** The results will be saved to an Excel file named `Efficient_Portfolio.xlsx`.

## Code Overview

The code is structured into several main parts:

- **Data Fetching:** Uses `yfinance` to download stock data and calculate log returns.
- **Data Processing:** Calculates the covariance matrix, correlation matrix, and efficient frontier.
- **GUI:** Created with `tkinter` for user interaction.
- **Excel Export:** Uses `openpyxl` to format and save data into Excel.

### Key Files

- `main.py`: The main script to run the application.
- `requirements.txt`: Lists required Python packages.

## Contributing

Contributions to the **Efficient Portfolio** project are welcome! Please follow these steps to contribute:

1. **Fork the Repository** and create a new branch.
2. **Make Changes** and commit your changes with descriptive messages.
3. **Push Changes** to your forked repository.
4. **Create a Pull Request** to the main repository.

If you have suggestions or improvements, please open an issue or submit a pull request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

For more detailed documentation and examples, please refer to the [Wiki](https://github.com/yourusername/efficient-portfolio/wiki).
