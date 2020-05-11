import timeit
import files_handling

def main():

    print("START - MAIN")
    start = timeit.default_timer()

    files_handling.main()

    stop = timeit.default_timer()
    print("END - MAIN")

    print('Time: ', stop - start)

    return True


if __name__ == '__main__': main()
