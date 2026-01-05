from src.canoe_can import CAN_CANoeConfig
from src.canoe_lin import LIN_CANoeConfig

# XML/XVP Configuration 
CANname = "Vehicle"
img_1 = "Logo-Stellantis.jpg"
img_2 = "Logo-Lear.png"

def main():
	# CAN 
	DBCfile = CAN_CANoeConfig()
	'''DBCfile.Print_Messages_Signals()'''
	DBCfile.Generate_EnVars()
	'''DBCfile.Generate_CAPL()'''
	DBCfile.Generate_CAPL_E2E()
	DBCfile.Generate_XVP(CANname, img_1, img_2)

	# LIN 
	LDFfile = LIN_CANoeConfig()
	'''LDFfile.Print_Messages_Parameters()'''
	'''LDFfile.Print_Signals_Parameters()'''
	LDFfile.Generate_EnVars()
	LDFfile.Generate_CAPL()


if __name__ == "__main__":
	main()
