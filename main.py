# Author: Selorm Kwami Dzakpasu

# Execute all scripts to generate the final combined generator interconnection queue database in .xlsx

exec(open("CAISO.py").read()) 

exec(open("ERCOT.py").read())

exec(open("MISO.py").read())

exec(open("PJM.py").read())

exec(open("SPP.py").read())

exec(open("NEISO.py").read())

exec(open("NYISO.py").read())

exec(open("Compiler.py").read())