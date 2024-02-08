import openpyxl as xl
from openpyxl.chart import BarChart, Reference
from cell_values_generator import generate_cell_values, calc_discounted_price, generate_bar_chart, generate_total_price

generate_cell_values(100)
calc_discounted_price()
generate_bar_chart()
generate_total_price()