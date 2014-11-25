# -*- coding: utf-8 -*-

import xlwt

def excel_style():
    alignmentVertical=xlwt.Alignment() #垂直居中
    alignmentVertical.vert=xlwt.Alignment.VERT_CENTER
    alignmentHorz=xlwt.Alignment() #水平居中
    alignmentHorz.horz=xlwt.Alignment.HORZ_CENTER
    styleVerticalHora = xlwt.Alignment() #垂直水平居中
    styleVerticalHora.vert=xlwt.Alignment.VERT_CENTER
    styleVerticalHora.horz=xlwt.Alignment.HORZ_CENTER
    alignmentWrap = xlwt.Alignment() #自动换行并垂直居中
    alignmentWrap.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    alignmentWrap.vert=xlwt.Alignment.VERT_CENTER
    #index=0 字体230大小 加粗 用于表头
    styleFont = xlwt.XFStyle()
    font = xlwt.Font()
    font.height=230
    font.bold=True
    styleFont.font=font
    #index=1 垂直居中
    styleVerticalCenter=xlwt.XFStyle()
    styleVerticalCenter.alignment=alignmentVertical
    #index=2 换行并垂直居中
    styleWordWrap = xlwt.XFStyle()
    styleWordWrap.alignment=alignmentWrap
    #index=3 时间格式化并垂直居中并靠左
    styleDate = xlwt.XFStyle()
    styleDate.num_format_str='YYYY-MM-DD hh:mm'
    styleDate.alignment=alignmentVertical
    styleDate.alignment.horz=xlwt.Alignment.HORZ_LEFT
    return {0:styleFont,1:styleVerticalCenter,2:styleWordWrap,3:styleDate}



def dict_out_excel(orders,columnList,sheetName="Sheet1"):
    styleDict=excel_style()
    file = xlwt.Workbook(encoding='utf-8')
    table = file.add_sheet(sheetName,cell_overwrite_ok=True)
    row=0

    #合计处理
    total={'colIndex':[],'operation':[],'val':[],'title':None}
    def total_dispose(i,valList,total=total):
        if i in total['colIndex']:
            index=total['colIndex'].index(i)
            op=total['operation'][index]
            if op=='+':
                for val in valList:
                    total['val'][index]+=val if type(val) in [int,float] else 0
    #写入表头
    for i,th in enumerate(columnList):
        header=th.get("header")
        if header:
            style=styleDict.get(header.get("style")) if type(header.get("style"))==int else header.get("style")
            table.write(0,i,header.get("title",""),style)
            if header.get("width",0)>0:
                table.col(i).width=header["width"]
            row=1

        elif th.get("total"):
            for coli in th["total"]["colIndex"]:
                total["colIndex"].append(coli)
                total["operation"].append("+")
                total["val"].append(0)
            total["title"]=th["total"]["title"] if th["total"].get("title") else ""
            columnList.remove(th)

    #写入正文
    for ord in orders:
        rowDataList=[]
        maxMergeNum=0 #记录最大合并数
        for ci,th in enumerate(columnList):
            field=th["field"]
            style=styleDict.get(th.get("style")) if type(th.get("style"))==int else th.get("style")
            valList=[]
            if type(field)==str:
                obj=ord
                for fie in field.split("."):
                    obj=obj.get(fie)
                    if obj==None:
                        break
                valList=[obj]
            elif type(field)==dict:
                model=field['model']
                if model in ["operation","item"]:
                    opVal=[]
                    vListNum=1
                    for col in field['col']:
                        colSplit=col.split(".")
                        if len(colSplit)>2:
                            raise Exception("Can't handle more than 2 levels of nesting(%s:%s)"%(model,col))
                        obj=ord.get(colSplit[0])
                        if obj in [None,[],{}]:
                            opVal.append(0)
                        elif type(obj)==dict:
                            obj=obj.get(colSplit[1])
                        elif type(obj)==list:
                            vListNum=len(obj)
                            vList=[]
                            for ob in obj:
                                vList.append(ob.get(colSplit[1]))
                            obj=vList
                        opVal.append(obj)

                    #填冲数组 方便运算
                    for i in range(len(opVal)):
                        if type(opVal[i])!=list:
                            opVal[i]=[opVal[i] for x in range(vListNum)]
                    opIndex=0   #标识运算符或连接符下标
                    if model=="operation":
                        while 1:
                            if len(opVal)==1:
                                break
                            nowOp=field["operation"][opIndex]
                            if nowOp=="+":
                                for i in range(vListNum):
                                    opVal[1][i]=opVal[0][i]+opVal[1][i]
                            elif nowOp=="-":
                                for i in range(vListNum):
                                    opVal[1][i]=opVal[0][i]-opVal[1][i]
                            elif nowOp=="*":
                                for i in range(vListNum):
                                    opVal[1][i]=opVal[0][i]*opVal[1][i]
                            elif nowOp in ["/","%"]:
                                for i in range(vListNum):
                                    if opVal[1][i]!=0:
                                        if nowOp=="/":
                                            opVal[1][i]=opVal[0][i]/opVal[1][i]
                                        else:
                                            opVal[1][i]=opVal[0][i]%opVal[1][i]
                                    else:
                                        opVal[1][i]=0
                            opVal.pop(0)
                            opIndex+=1
                    elif model=="item":
                        join=field.get("join",["" for x in range(vListNum)])
                        while 1:
                            if len(opVal)==1:
                                break
                            j=join[opIndex]
                            for i in range(vListNum):
                                opVal[1][i]="%s%s%s"%(opVal[0][i],j,opVal[1][i])
                            opVal.pop(0)
                            opIndex+=1
                    valList=[v for v in opVal[0]]

                elif model=="join":
                    opVal=[]
                    for col in field["col"]:
                        obj=ord
                        for fie in col.split("."):
                            obj=obj.get(fie)
                        opVal.append(obj)
                    join=field.get("join")
                    if join:
                        for i in range(len(opVal)-1):
                            if i>len(join)-1:
                                j=join[-1]
                            else:
                                j=join[i]
                            opVal[1]="%s%s%s"%(opVal[0],j,opVal[1])
                            opVal.pop(0)
                    else:
                        opVal[0]="".join(opVal)
                    valList=[opVal[0]]

            if maxMergeNum<len(valList):
                rowDataList.insert(0,{"merge":len(valList),"col":ci,"val":valList,"style":style})
                maxMergeNum=len(valList)
            else:
                rowDataList.append({"merge":len(valList),"col":ci,"val":valList,"style":style})
            total_dispose(ci,valList)

        mergeNum=rowDataList[0]["merge"]
        for rd in rowDataList:
            if rd["merge"]<mergeNum:
                rd["val"]=[str(v) if v!=None else "" for v in rd["val"]]
                table.write_merge(row,row+mergeNum-1,rd["col"],rd["col"],"\n".join(rd["val"]),rd["style"])
            else:
                for i,v in enumerate(rd["val"]):
                    v=str(v) if v!=None else ""
                    table.write(row+i,rd["col"],v,rd["style"])
        row+=mergeNum
    if total["title"]!=None:
        table.write(row,0,total['title'])
        for i in total['colIndex']:
            index=total['colIndex'].index(i)
            v=total['val'][index]
            table.write(row,i,v)

    return file


#测试
columnList=[
    {"field":"status.name","style":1,"header":{"title":"A","width":3000,"style":0}},
    {"field":{"model":"item","field":"items","col":["a","items.goods_num","b"],"join":["*","/"]},"style":2,"header":{"title":"B","width":10000,"style":0}},
    {"field":{"model":"operation","col":["a","items.goods_num","b"],"operation":["*","/"]},"style":2,"header":{"title":"C","width":10000,"style":0}},
    {"field":"a","style":1,"header":{"title":"D","width":3000,"style":0}},
    {"field":{"model":"join","col":["a","b","status.name"],"join":["-","*"]},"style":1,"header":{"title":"E","width":3000,"style":0}},
    {"total":{"title":"合计","colIndex":[2,3]}}
]

orders=[
    {"a":1,"status":{"name":"a"},"b":2,"items":[{"goods_num":3,"price":5},{"goods_num":4,"price":6}]},
    {"a":2,"status":{"name":"b"},"b":3,"items":[{"goods_num":4,"price":6},{"goods_num":5,"price":7}]}
]
file=dict_out_excel(orders,columnList)
import os
i=1
while 1:
    if os.path.isfile("e:/%s.xls"%i):
        i+=1
    else:
        break
file.save("e:/%s.xls"%i)
