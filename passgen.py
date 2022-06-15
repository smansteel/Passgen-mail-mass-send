from random import randint
import pandas as pd
import numpy as np


charlist = ['a', 'z', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p', 'q', 's', 'd', 'f', 'g', 'h', 'j', 'k', 'm', 'w', 'x',
            'c', 'v', 'b', 'n', '?', '.', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '-', '_', 'é', '&', 'A',
            'Z', 'E', 'R', 'T', 'Y', 'U', 'O', 'P', 'Q', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'W', 'X', 'C',
            'V', 'B', 'N']

num = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']
maj = ['A', 'Z', 'E', 'R', 'T', 'Y', 'U', 'O', 'P', 'Q', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'W', 'X', 'C',
       'V', 'B', 'N']
low = ['a', 'z', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p', 'q', 's', 'd', 'f', 'g', 'h', 'j', 'k', 'm', 'w', 'x',
       'c', 'v', 'b', 'n']
spe = ['?', '.', '-', '_', 'é', '&']
def passgen():
    psswd = ""
    for i in range(0, 9):
        psswd += charlist[randint(0, len(charlist)-1)]

    if (any(x in psswd for x in num)):
        pass
    else:
        psswd += num[randint(0, len(num)-1)]
    if (any(x in psswd for x in maj)):
        pass
    else:
        psswd += maj[randint(0, len(maj)-1)]
    if (any(x in psswd for x in low)):
        pass
    else:
        psswd += low[randint(0, len(low)-1)]
    if (any(x in psswd for x in spe)):
        pass
    else:
        psswd += spe[randint(0, len(spe)-1)]


    return (psswd)


print(passgen())


file_name =  """Input file path"""# path to file + file name
sheet =  "Feuil1"# sheet name or sheet number or list of sheet numbers and names

df = pd.ExcelFile(file_name).parse(sheet)

print(len(df))
pwdlist = []
for i in range(0,len(df)):
    pwdlist.append(passgen())
df["Newpwd"] = pwdlist
df.to_excel("output file path")
print(df)

