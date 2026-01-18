# TOPSIS Implementation in Python

**Course:** UCS654 - Predictive Analytics using Statistics  
**Assignment:** Assignment-1 (TOPSIS)  
**Author:** Bhavya  
**Roll Number:** 102303560  

---

## About the Project

This repository contains a Python implementation of the  
**TOPSIS (Technique for Order of Preference by Similarity to Ideal Solution)** method.

TOPSIS is a **multi-criteria decision-making (MCDM)** technique used to rank multiple
alternatives based on their distance from the ideal best and ideal worst solutions.

---

## Installation - USER MANUAL
1. Topsis-bhavya-102303560 requires Python3 to run.
2. Other dependencies that come installed with this package are :-
    - pandas
    - numpy
3. Package listed on PyPI:- https://pypi.org/project/Topsis-bhavya-102303560/
4. Use the following command to install this package:-
    ```bash
    pip install Topsis-bhavya-102303560

---

## Usage
Run the following command in command prompt:
```bash
topsis <inputFile> <weights> <impacts> <outputFile>
```
Example:
```bash
topsis sample.csv "1,1,1,1" "+,+,-,+" result.csv
```
