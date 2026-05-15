import csv
from itertools import combinations
a=["BTD", "DPT", "EMS", "HKK", "ICK", "SZK"]
b=["AHEM", "ANZE", "AZEC", "BCME", "CLTG-F", "DGAH", "DGAH-F", "DGCY", "DGCY-F", 
        "DGQG", "DGQG-F", "HKYQ-F", "HQBS", "HQDZ", "HQDZ-F", "HT", "JSCC", "JSCC-F", 
        "JYZZ", "JYZZ-F", "KLEM", "KLKT", "KLKT-F", "KPMI", "KPMI-F", "KTRN", "MNST", 
        "MOLT", "NGSB", "NGSB-F", "NJVtest-F", "OLBS", "OLEM", "OLTT", "OLTT-F", "ONET", 
        "ONET-F", "OVM", "PIEM", "PIEM-F", "QGEM", "SCK", "SMAT", "SMTT", "SMTT-F", "SPIL", 
        "SPILSZ", "TF", "TFME", "THQG", "THQG-F", "Vtest-F", "WINSTEK", "WINSTEK-F", "WXHN-F", 
        "XCEM", "XJCS-F", "XMXC", "XMXC-F", "YDFT", "YDFT-F", "YNAH", "YNAH-F", "YZKJ"]
ans=[]
for i in a:
    for j in b:
        ans.append(i+j)
print(ans)

