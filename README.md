# Wind ESG Data Automation Tool

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![WindPy](https://img.shields.io/badge/WindPy-API-green)
![License](https://img.shields.io/badge/License-MIT-orange)

## Project Description
Automated tool for downloading listed companies' ESG data from Wind Financial Terminal, including:
- Comprehensive ratings (`esg_rating_wind`)
- E/S/G component scores (`esg_escore`/`esg_sscore`/`esg_gscore`)
- Carbon emission data (`esg_ghg*` series)
- International agency ratings (MSCI/FTSE Russell)

## Quick Start

### Prerequisites
1. Install Wind Official Python API (requires Wind account)
   ```bash
   pip install WindPy
