#### 说明：  
本程序用于从基金列表（存于excel文件中）中筛选出满足要求的基金。  

运行环境：  
python3.   
wlrd，wlwt(位于thirdparty中)   
程序运行需要调用 `my_package` 中的代码，所以必须让程序能找到它。例如`my_package`被下载到~/python_study目录中，就要执行：
export PYTHONPATH=~/python_study   
基金列表从晨星基金网手动copy到excel文件中，拷贝时分别按前5年业绩，3年业绩排序，粘贴到excel中后要手动检查一下：在基金收益率的cell中不能有‘-’ 等非数字  

命令执行方法：  
参数1：只考虑3年业绩情况，可选择5  
参数2：基金业绩表格  
参数3：基金规模表格  
```
python3 find_best_stock.py 3 ./data/201905/debt.xlsx ./data/201905/debt_scale.xlsx
```