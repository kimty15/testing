from spire.xls import *
from spire.xls.common import *
def excel2image(path:str, save_path:str):
    workbook = Workbook()
    workbook.LoadFromFile(path)
    worksheet = workbook.Worksheets[0]
    # worksheet.to_image().save(path)
    image = worksheet.ToImage(worksheet.FirstRow, worksheet.FirstColumn, worksheet.LastRow, worksheet.LastColumn)
    image.Save(save_path)
    workbook.Dispose()

path = r"/workspaces/testing/06J0690_SI_VGM_SP_2EF56.xls"
save_path = "demo.png"

excel2image(path, save_path)