#!/bin/sh
export PYTHONPATH=~/python_study
echo hybrid 3 year
python3 find_best_stock.py 3 ./data/201802/hybrid_201802_3y.xlsx
echo hybrid 5 year
python3 find_best_stock.py 5 ./data/201802/hybrid_201802_5y.xlsx
echo stock 3 year
python3 find_best_stock.py 3 ./data/201802/stock_201802_3y.xlsx
echo stock 5 year
python3 find_best_stock.py 5 ./data/201802/stock_201802_5y.xlsx
echo debt 3 year
python3 find_best_stock.py 3 ./data/201802/debt_201802_3y.xlsx
echo debt 5 year
python3 find_best_stock.py 5 ./data/201802/debt_201802_5y.xlsx
