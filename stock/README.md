#### 说明：  
本程序用于从基金列表（存于excel文件中）中筛选出满足要求的基金。  

运行环境：  
python3.   
wlrd，wlwt(位于thirdparty中), 要安装它们，安装方法是：  
将他们拷贝到Python的安装目录下的`Lib\site-packages`目录中。然后进入该目录执行：  
`
 python -m pip install xlwt-1.3.0-py2.py3-none-any.whl 
 `   
 ` 
 python -m pip install xlrd-1.1.0-py2.py3-none-any.whl
 `   
 
程序运行需要调用 `my_package` 中的代码，所以必须让程序能找到它。
find_best_stock.py中已经指定了搜索路径
`
sys.path.append(r'../')
`  
基金列表从晨星基金网手动copy到excel文件中，拷贝时分别按前5年业绩，3年业绩排序，粘贴到excel中后要手动检查一下：在基金收益率的cell中不能有‘-’ 等非数字  

命令执行方法：  
参数1：只考虑3年业绩情况，可选择5  
参数2：基金业绩表格  
参数3：基金规模表格  
```
python3 find_best_stock.py 3 ./data/201905/debt.xlsx ./data/201905/debt_scale.xlsx
```
