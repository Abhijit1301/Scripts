# Python program to demonstrate
# command line arguments


import getopt, sys


def parse_using_getopt():
    # Remove 1st argument from the
    # list of command line arguments
    argumentList = sys.argv[1:]

    # Options
    options = "hmo:"

    # Long options
    long_options = ["Help", "My_file", "Output="]

    try:
        # Parsing argument
        arguments, values = getopt.getopt(argumentList, options, long_options)
        
        print("arguments: ")
        print(arguments)
        print("values: ")
        print(values)
        
        # checking each argument
        for currentArgument, currentValue in arguments:

            if currentArgument in ("-h", "--Help"):
                print ("Displaying Help")
                
            elif currentArgument in ("-m", "--My_file"):
                print ("Displaying file_name:", sys.argv[0])
                
            elif currentArgument in ("-o", "--Output"):
                print (("Enabling special output mode (% s)") % (currentValue))
                
    except getopt.error as err:
        # output error, and return with an error code
        print (str(err))

def parse_using_argparse_module():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--Initialize', action="store_true", help='Initialize the databases.')
    parser.add_argument('-c', '--Cleanup', action="store_true", help='Cleanup the databases.')
    args = parser.parse_args()
    if not (args.Initialize or args.Cleanup):
        parser.error("At-least one argument is needed")
    
    return args

if (__name__ == "__main__"):
    print("This was invoked directly...............................")
    args = parse_using_argparse_module()
    print(args)
    
else:
    print("This is imported as a module")