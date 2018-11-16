import xlrd
import os
def getexcels():
    PathArr = []
    path = os.getcwd()
    n = 0
    f = os.listdir(path)
    for i in f:
        oldname=str(path)+'\\'+str(f[n])
        if oldname[-5:] == '.xlsx':
            if '~$' in oldname:
                continue
            else:
                PathArr.append(oldname)
        n = n + 1
    return PathArr
def FindinExcel(EcLn,StrLf):
    EcL= xlrd.open_workbook(EcLn)
    ShNarr = EcL.sheet_names()
    ShArr = EcL.sheets()
    RtArr = []
    Rtstr = ''
    for ShN in ShNarr:
        Sh = EcL.sheet_by_name(ShN)
        nrows = Sh.nrows
        ncols = Sh.ncols
        for i in range(nrows): 
            row = Sh.row_values(i)
            cl = 0
            for n in row:
                cl = cl + 1
                if StrLf in str(n):
                    Rtstr = str("工作表："+str(EcLn)+"工作簿："+str(ShN)+"行数："+str(i+1)+"列数："+str(cl))
                    RtArr.append(Rtstr)
    return RtArr
def main():
    LkF = input("输入你想查找的字符\n")
    print("查找中...")
    LocArr = getexcels()
    for i in LocArr:
        ans = FindinExcel(i,LkF)
        for t in ans:
            print(t)
    print("查找完毕!")
    input("输入任意字符退出")
main()
            
