from canoe_autogen_can import CAN_CANoeConfig

# XML/XVP Configuration 
CANname = "Vehicle"
img_1 = "Logo-Stellantis.jpg"
img_2 = "Logo-Lear.png"

def main():
	DBCfile = CAN_CANoeConfig()
	DBCfile.find_input_paths()
	DBCfile.Print_Messages_Signals()
	DBCfile.Generate_EnVars()
	# DBCfile.Generate_CAPL()
	DBCfile.Generate_CAPL_E2E()
	DBCfile.Generate_XMLcode(CANname, img_1, img_2)

if __name__ == "__main__":
	main()
