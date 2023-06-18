# coding=utf-8
import openpyxl

# 创建文件
workbook = openpyxl.Workbook()

# 找到当前清单
worksheet = workbook.active

# 宇视Unigulf系列室内小间距LED系统报价清单
# 写入内容
worksheet['A1'] = u'宇视Unigulf系列室内小间距LED系统报价清单'
worksheet['A1'].alignment = openpyxl.styles.Alignment(horizontal='center')
worksheet['A1'].font = openpyxl.styles.Font(name='Microsoft YaHei', size=18)
worksheet.row_dimensions[1].height = 25

# 加入图标
img = openpyxl.drawing.image.Image('unigulf3.png')
img.height = int(img.height/1.8)
img.width = int(img.width/1.8)
worksheet.add_image(img, 'A1')

# 合并单元格
worksheet.merge_cells('A1:J1')

# 序号	设备名称	产品型号	技术参数	品牌	 数量	单位	单价	总价	备注
l = "序号	设备名称	产品型号	技术参数	品牌	 数量	单位	单价	总价	备注".split()
width = (4,16,10,40,9,5,7,8,10,16)
for i in range(10):
    worksheet[chr(i+65)+'2']=l[i]
    worksheet[chr(i + 65) + '2'].font =  openpyxl.styles.Font(name='Microsoft YaHei', size=11, bold=True)
    worksheet[chr(i + 65) + '2'].alignment = openpyxl.styles.Alignment(horizontal='center')
    worksheet.column_dimensions[chr(i + 65) ].width = width[i]

# 1,2,3....
products_dict = {'LED':None,
                 'ReceivingCard':None,
                 'PowerSupply':None,
                 'SendingCard':None,
                 'DistributionBox':None,
                 'Installation':None,
                 'Software':None}





# 保存内容
workbook.save('test.xlsx')
