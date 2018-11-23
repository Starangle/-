from win32com.client import constants, gencache, Dispatch
gencache.EnsureDispatch('Word.Application')
wordapp = Dispatch('Word.Application')

# 参数区，每次运行前需要修改这些数据
workDir = r'C:/Users/binsu/Desktop/Invitation/贾老师组邀请信生成模板及名单/'    #工作目录
templatefileName = r'template.docx' #邀请信模板
targetList = [r'Haozhe WU',r'Haozhe WANG']    #邀请的对象们
templateStr = r'(Name)' #要被替换的字符串
# 参数区结束

for item in targetList:
    doc = wordapp.Documents.Open(workDir+templatefileName)
    wordapp.Selection.Find.ClearFormatting()
    wordapp.Selection.Find.Replacement.ClearFormatting()
    wordapp.Selection.Find.Execute(
        templateStr, False, False, False, False, False, True, 1, True, item, 2)
    doc.ExportAsFixedFormat(workDir+item.replace(' ','')+'.pdf', constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup)
    doc.Close(constants.wdDoNotSaveChanges)
wordapp.Quit()

###实现参考
# https://www.cnblogs.com/baiboy/p/7251484.html
# https://stackoverflow.com/questions/28264548/how-to-use-win32com-client-constants-with-ms-word
# https://docs.microsoft.com/zh-cn/office/vba/api/word.selection.exportasfixedformat
