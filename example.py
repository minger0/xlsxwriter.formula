import xlsxwriter
import xlsxwriter_formula as x

dom_size = ["S","M","L"]
dom_food = ["pizza","pasta","soup"]
dom_extra = ["mushroom","salami","bacon"]

tableinput = {
	"name"   : "input"
, "anchor" : [0, 0]
, "row"    : dom_size
, "col"    : dom_food
, "val"    : lambda size, food : food + ' ' + size
}

tablederive = {
	"name"   : "derive"
, "anchor" : [0, 5]
, "row"    : dom_size
, "col"    : dom_extra
, "val"    : lambda size, extra : f'=CONCATENATE({x.vref("input",size,"pizza")}," - ",{x.cref("derive",extra)})'
}

book = xlsxwriter.Workbook('example.xlsx')
sheet = book.add_worksheet('overview')

x.View(sheet,tableinput).populate()
x.View(sheet,tablederive).populate()

book.close()

