from os import listdir, path, chdir, getcwd
from sys import exit
from struct import unpack, pack

CHECKSUM_LINE = 7
CHECKSUM_WORD = 0x3F
CHECKSUM_VALUE = 0xBABA

def calc_unsinged_short_csum(sum):

    #transform to unsigned integer
    sum += 2**32

    #add checksum value
    sum += CHECKSUM_VALUE

    return (sum & 0xFFFF)

def update_files_in_dir():

    #update binary file
    files = listdir()
    binary_file_name = None

    for file in files:
        if file.endswith(".bin"):
            print(f"Found binary file: {file}")
            binary_file_name = file
            break

    if binary_file_name == None:
        exit("Couldn't find any binary file")

    sum = 0
    with open(binary_file_name, "r+b") as word:
        for offset in range(0, CHECKSUM_WORD):
            word_value = hex(unpack('<H', word.read(2))[0])
            sum += ((int(word_value, 16)))

        checksum_value = calc_unsinged_short_csum(-sum)
        print(f"Calculated checksum from binary file: {hex(checksum_value)}")

        csum_value = pack('<H', checksum_value)
        word.write(csum_value)

    #update txt file
    txt_file_name = None
    for file in files:
        if file.endswith(".txt"):
            txt_file_name = file
            print(f"Found text file: {file}")
            break
    if txt_file_name == None:
        exit("Couldn't find any text file")


    #calculate the correct checksum from the txt file as a sanity check
    sum = 0
    with open(txt_file_name, "r+") as txt_file:
        finished = False
        word_count = 0

        for line in txt_file:
            for word in line.split():
                if word_count == CHECKSUM_WORD:
                    txt_checksum_value = calc_unsinged_short_csum(-sum)
                    print(f"Calculated checksum value from text file: {hex(txt_checksum_value)}")
                    finished = True
                    break

                word_value = "0x" + word
                sum += ((int(word_value, 16)))
                word_count += 1
            if finished:
                break

    if checksum_value != txt_checksum_value:
        exit("txt and binary checksum values differ")

    checksum_value = hex(checksum_value)[2:]

    # Replace the checksum in the text file
    with open(txt_file_name, "r") as txt_file:
        data = txt_file.readlines()
        checksum_line = data[CHECKSUM_LINE]
        checksum_line = checksum_line.split()[:-1]
        checksum_line.append(checksum_value.upper())
        data[CHECKSUM_LINE] = " ".join(checksum_line)

        #add newline to the end
        data[CHECKSUM_LINE] = f"{data[CHECKSUM_LINE]}\n"

    with open(txt_file_name, "w") as txt_file:
        txt_file.writelines(data)

dir_list = [name for name in listdir() if path.isdir(name)]
print(f"Current working directory: {getcwd()}\n")
print(f"Directories found: {dir_list}")

for dir in dir_list:
    print(f"\nUpdating dir: {dir}")
    chdir(dir)
    update_files_in_dir()
    chdir("..")