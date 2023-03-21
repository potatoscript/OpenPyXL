# Python and Excel Programming with OpenPyXL

- Create working dir
  - `mkdir /c/excel-python`
- Create virtual environment
  - `python -m venv virt` where `virt` is the name of your virtual environment
- to list the virtual environment type `ls`
- to turn on your virtual environment `source virt/Scripts/activate`
- to turn it off `deactivate`
- Install OpenPyXL
  - `pip freeze` this is to freeze everything and show us what is been installed
  - `pip install openpyxl` this will install into our virtual environment only
- import open piexcel into our code

  - open your project file in your vs code and type the following lines on the top of the file

    ```python
      from openpyxl.wworkbook import Workbook
      from openpyxl import load_workbook

      #Create a workbook object
      wb = Workbook()
      # or load existing spredsheet
      wb = load_workbook('hello.xlsx')
      # Create a worksheet object
      ws = wb.active
      # Print something from our spreadsheet
      name = ws['A2'].value
      color = ws['B2'].value
      print(f'{name}: {color}')
    ```
