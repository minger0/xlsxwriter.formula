''' a layer on top of xlsxwriter to support excel formula creation, a lightweight version '''
# BSD license 2.0 (3-clause)
# Copyright 2016, Gergely Mincsovics

# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:

		# 1. Redistributions of source code must retain the above copyright notice,
		# this list of conditions and the following disclaimer.
		# 2. Redistributions in binary form must reproduce the above copyright notice,
		# this list of conditions and the following disclaimer in the documentation and/or
		# other materials provided with the distribution.
		# 3. Neither the name of the copyright holder nor the names of its contributors may be
		# used to endorse or promote products derived from this software without specific
		# prior written permission.

# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR
# IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
# FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
# CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
# DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER
# IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT
# OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

from xlsxwriter.utility import xl_rowcol_to_cell

v = {}

def vref(view, row, col):
	''' reference to a view at its row and column '''
	return xl_rowcol_to_cell( v[view].anchor[0] + 1 + v[view].rowheader.index(row)
		, v[view].anchor[1] + 1 + v[view].colheader.index(col))

def rref(view, row):
	''' row reference to a view at a given row '''
	return xl_rowcol_to_cell( v[view].anchor[0] + 1 + v[view].rowheader.index(row)
		, v[view].anchor[1], col_abs = True)

def cref(view, col):
	''' column reference to a view at a given column '''
	return xl_rowcol_to_cell( v[view].anchor[0]
		, v[view].anchor[1] + 1 + v[view].colheader.index(col), row_abs = True)

class View():
	''' a view is a table, a part of a work sheet '''

	def __init__(self, sheet, viewdef, val=None):
		self.sheet = sheet
		self.anchor     = viewdef["anchor"]
		self.name       = viewdef["name"]
		self.rowheader  = viewdef["row"]
		self.colheader  = viewdef["col"]
		self.funcvalue  = viewdef.get("val")        # optional
		if self.funcvalue is None:
			self.funcvalue = val
		self.funcformat = viewdef.get("funcformat") # optional
		self.rowcount   = len(self.rowheader)
		self.colcount   = len(self.colheader)
		self.col = 0
		self.row = 0
		v[self.name] = self

	def __iter__(self):
		self.col = 0
		self.row = 0
		return self

	def __next__(self):
		if self.row > self.rowcount:
			raise StopIteration
		if self.col == 0 and self.row == 0:
			retval = self.name
		elif self.row == 0:
			retval = self.colheader[self.col-1]
		elif self.col == 0:
			retval = self.rowheader[self.row-1]
		else:
			retval = self.funcvalue(self.rowheader[self.row-1], self.colheader[self.col-1])

		print(self.name,"(",self.rowheader[self.row-1],",",self.colheader[self.col-1],") "
			+ "coord rel [",self.row,",",self.col,"] coord abs [",self.sheetrow(),self.sheetcol(),"] = ",retval)
		retfield = {"row": self.sheetrow(), "col": self.sheetcol(), "val": retval}

		self.nextfield()

		return retfield

	def sheetrow(self):
		return self.anchor[0]+self.row

	def sheetcol(self):
		return self.anchor[1]+self.col

	def nextfield(self):
		self.col+=1
		if self.col > self.colcount:
			self.row+=1
			self.col=0

	def populate(self):
		for field in iter(self):
			self.sheet.write(field["row"], field["col"], field["val"])

