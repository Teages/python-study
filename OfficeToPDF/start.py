# -*-coding:utf-8-*-
import os
import shutil
#import requests


output_location = "pdf";
input_location = "class";

support_file = "doc,dot,docx,dotx,docm,dotm,rtf,wpd,xls,xlsx,xlsm,xlsb,xlt,xltx,xltm,csv,ppt,pptx,pptm,pps,ppsx,ppsm,pot,potx,potm";

#if not os.path.exists("OfficeToPDF.exe") :
#    with open("OfficeToPDF.exe",'wb') as main_program :
#        main_program.write(requests.get("https://github.com/cognidox/OfficeToPDF/releases/download/v1.9.0.2/OfficeToPDF.exe").content)

#os.remove("in_log.txt");
#os.remove("out_log.txt");
for root,dirs,files in os.walk(".\\"+ input_location +"\\"):
    for file in files:
        #print os.path.join(root,file).decode('gbk').encode('utf-8');
        if not os.path.exists( os.path.join(root,file).replace(file,"").replace(".\\",".\\" + output_location + "\\") ):
            os.makedirs(os.path.join(root,file).replace(file,"").replace(".\\",".\\" + output_location + "\\"));
        if os.path.exists( os.path.join(root,file).replace(".\\",".\\" + output_location + "\\").replace( os.path.splitext(file)[1] ,".pdf" ) ):
            continue;
        if os.path.splitext(file)[1].replace('.','') in support_file :
            #os.system("echo \"" + os.path.join(root,file) + "\" >> in_log.txt");
            #os.system("echo \"" + os.path.join(root,file).replace(".\\",".\\" + output_location + "\\") + "\" >> out_log.txt");
            print(".\\OfficeToPDF.exe \"" + os.path.join(root,file) + "\" \"" + os.path.join(root,file).replace(".\\",".\\" + output_location + "\\").replace( os.path.splitext(file)[1] ,".pdf" ) + "\"");
            os.system("echo .\\OfficeToPDF.exe \"" + os.path.join(root,file) + "\" \"" + os.path.join(root,file).replace(".\\",".\\" + output_location + "\\").replace( os.path.splitext(file)[1] ,".pdf" ) + "\">> log.txt");
            os.system(".\\OfficeToPDF.exe \"" + os.path.join(root,file) + "\" \"" + os.path.join(root,file).replace(".\\",".\\" + output_location + "\\").replace( os.path.splitext(file)[1] ,".pdf" ) + "\"");
        else :
            shutil.copy(os.path.join(root,file) , os.path.join(root,file).replace(".\\",".\\" + output_location + "\\"));
