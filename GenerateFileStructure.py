import os
import shutil

BaseDir = "C:\\Users\\Weston\\Desktop\\AdubsAuto\\QAs"
TemplatePath = "C:\\Users\\Weston\\Desktop\\AdubsAuto\\template.xlsx"
NumberOfBatches = 4


Agents = [
    "Amy",
    "Fred",
    "Tim"
]


"""
Agents = [
    "Amy",
    "Fred",
    "Tim"
]
"""

Months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sept",
    "Oct",
    "Nov",
    "Dec",
]

def get_path(m, a, i):
    output = ""
    output += '\\'
    output += m
    output += '\\'
    output += a
    output += '\\'
    output += str(i)
    
    return output
    
def get_file_name(a, mi, batch):
    output = ""
    output += a
    output += ' '
    output += Months[mi]
    output += ' '
    output += str(batch)
    output += ".xlsx"
    
    return output
    
def build_months():
    for i in range(len(Months)):
        month = str(i+1) + ' ' + Months[i]
        curpath = BaseDir
        curpath += '\\'
        curpath += month
        
        try:
            os.mkdir(curpath)
            build_agents(curpath, i)
        except FileExistsError:
            build_agents(curpath, i)

def build_agents(base, month):  
    for agent in Agents:
        curpath = base
        curpath += '\\'
        curpath += agent
        
        try:
            os.mkdir(curpath)
            for i in range(1, NumberOfBatches + 1):
                final_path = curpath + '\\' + str(i)
                os.mkdir(final_path)
                dst = final_path + '\\' + get_file_name(agent, month, i)
                print(dst)
                shutil.copyfile(TemplatePath, dst)
        except FileExistsError:
            pass

def main():
    build_months()
                
main()