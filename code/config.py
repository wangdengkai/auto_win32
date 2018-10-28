#目录列表,可以处理多个目录
dir_list=[r'C:\Users\Administrator\Desktop\new_new']
#正文更换字段(old_str,new_str)
text_list =[
    ('old_str','new_str'),
    ('质量','品质')
]
#页mei
page_header_list=[
    ('old_str','new_str'),
    ('质量','品质')
]
#页脚
page_footer_list=[
    ('old_str','new_str'),
    ('27','31')
]
#filename
file_name=('2015','2018')


template_file=r'C:\Users\Administrator\Desktop\new_new\template.docx'
dest_file = r'C:\Users\Administrator\Desktop\new_new\2018.docx'

temp_dir=r'C:\Users\Administrator\Desktop'
word_before=r'以下为新增单选题目7个：'


context_mark=(
    '{{title}}','{{ans}}','{{A}}','{{B}}','{{C}}','{{D}}'
)
data_list=[
    (
        '以下关于品质活动说法错误的是','B','品质活动的母的是提升项目交付品质','品质活动是QA实施','品质活动的实施需要根据项目特点','品质活动需要项目组全员的共同关注'
    ),
    (
        '关于评审流程，下述说法错误的是','C','PM需要制定评审计划，确认组织者和评审人员等关键信息','介绍会议根据评审专家对评审对象的熟悉程度可以进行裁剪','准备阶段（预审阶段）是可以裁减的','第三小时会议的目的是讨论问题的解决方案 '
    )

]