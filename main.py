from xlsxfile import XLSXFile
from datamanipulator import DataManipulator
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilenames, askdirectory
import os, argparse


# Create the parser
my_parser = argparse.ArgumentParser(prog="GPC Normaliser",
                                    usage="%(prog)s --mode [arg1] --ranges [start_time] [stop_time] , [start_time] [stop_time], [start_time] [stop_time] ...",
                                    description='GDB GPC Normalisation Tool v1.0 by KS')

# Add the arguments
my_parser.add_argument("-m", "--mode", action='store', type=str, required=True, help="usage: -m [arg] where [arg] can be: kientic -> kinetic normalisation of the set only. height -> height normalisation of the set only. both -> both kinetic and height normlisation will be performed on the set.")
my_parser.add_argument("-r", "--ranges", action='store', nargs="+", type=str, required=True, help="usage: -r [arg...] where [arg...] are specified as follows: [start_time_1 stop_time_1 start_time_2 stop_time_2 start_time_3 stop_time_3]")
my_parser.add_argument("-c", "--combination", action="store_true", default=False, required=False, help="usage: -c true -> (defaults to false if not called) overrides the format of the input of the of start and stop times in --ranges argument: [start_time_1 start_time_2 ... ! stop_time_1 stop_time_2 ...] The colletion does not have to have equal number of start and stop times, the combination of all will be generated.")


Tk().withdraw()
CWD = os.path.dirname(__file__)
DAT = DataManipulator()
ARGS = my_parser.parse_args()
print("\n\n")




def main(args):
    mode = DAT.check_mode(arg=args.mode)
    ranges = DAT.check_ranges_input(args=args.ranges, combination=args.combination)

    paths = askopenfilenames(title="Select your data set:")
    if len(paths) == 0:
        exit("No paths were provided... Try again wiht some files...")
    if len(ranges) < 1:
        exit("Could not have parsed your time ranges, please double check your input.")
    files = [XLSXFile(path=path) for path in paths]
    DAT.number_of_files = len(files)
    print("\n\nExtraction and normalisation in progress...")

    for range in ranges:
        for file in files:
            if mode == "height":
                DAT.height_normalise(file=file, start=range[0], stop=range[1])
            elif mode == "both":
                DAT.height_normalise(file=file, start=range[0], stop=range[1])
        
        if mode != "height":
            DAT.kinetic_normalise(files=files, start=range[0], stop=range[1])

    print("Normalisation complete...")

    while True:
        dir_name = askdirectory(title="Save the results in directory:")
        if dir_name == "" or dir_name == None:
            answer = messagebox.askyesno(title="Try again?", message="You did not pick a directory to save your results... Try again?")
            if answer == False:
                exit("Your nomalisation data was not saved.")
        else:
            break

    print("Saving results...\n\n")
            
    if mode == "height":
        DAT.save_height_data_to_file(save_directory=dir_name, files=files)
    elif mode == "kinetic":
        DAT.save_kinetics_data_to_file(save_directory=dir_name)
    else:
        DAT.save_height_data_to_file(save_directory=dir_name, files=files)
        DAT.save_kinetics_data_to_file(save_directory=dir_name)




if __name__ == "__main__":
    main(args=ARGS)