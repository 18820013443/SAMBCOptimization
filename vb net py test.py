import time
import sys

class Test:

    def __init__(self) -> None:
        
        pass

    @staticmethod
    def test():
        try:
            # Your Python script code here
            result = 1 / 0
            print(result)  # This won't execute due to the error
        except Exception as e:
            # Redirect the error to standard output
            print("Error:", str(e), file=sys.stderr)
            # print("Error:", str(e), file=sys.stdout)
            # sys.stdout.write('Error:%s\n'%(str(e)))
            sys.exit(1)
            pass



if __name__ == '__main__':
    Test.test()