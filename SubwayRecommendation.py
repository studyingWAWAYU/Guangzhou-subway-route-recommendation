# -*- coding:utf-8 -*-
import openpyxl
import networkx as nx
import matplotlib.pyplot as plt
import tkinter as tk
from PIL import Image, ImageTk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

INF = float('inf')  # 定义正无穷

def get_data():  # 从Excel文件中获取数据并存入列表
    # 打开Excel文件并获取工作表
    workbook = openpyxl.load_workbook('Info.xlsx')
    worksheet = workbook.active

    names = []
    route = []
    adj = []
    distance = []
    time = []
    position = []

    # 将文件中数据放入列表
    for row in worksheet.iter_rows(min_row=2, values_only=True,max_row=133):
        names.append(row[0])
        eachroute = [int(num) for num in str(row[1]).split(",")]
        route.append(eachroute)
        eachadj=[str(adjname) for adjname in row[2].split("、")]
        adj.append(eachadj)
        eachdistance=[float(num) for num in str(row[3]).split(",")]
        distance.append(eachdistance)
        eachtime=[int(num) for num in str(row[4]).split(",")]
        time.append(eachtime)
        position.append([int(row[5]),int(row[6])])

    return names, route, adj, distance, time, position

def create_adj_matrix(names): # 生成邻接矩阵
    matrix = [[0 if i == j else INF for j in range(len(names))] for i in range(len(names))]
    for i in range(len(adj)): #按数据加边和权重
        row = i
        for j in range(len(adj[i])):
            col = names.index(adj[i][j])
            matrix[row][col] = [distance[i][j],time[i][j],1]
            matrix[col][row] = [distance[i][j],time[i][j],1]  #因为地铁双向，所以是无向图
    return matrix

def showGraphTotal(matrix,fig): #构建地铁线路的总图
    G = nx.DiGraph()
    # 添加节点并设置节点属性
    for i in range(len(matrix)):
        if route[i][0] == 1:
            rcolor = "#f9eec8"
        elif route[i][0] == 2:
            rcolor = "#d7dee7"
        elif route[i][0] == 3:
            rcolor = "#fae8db"
        elif route[i][0] == 4:
            rcolor = "#d6e3d8"
        elif route[i][0] == 5:
            rcolor = "#f8eaec"
        elif route[i][0] == 6:
            rcolor = "#f0eaed"
        G.add_node(i, color=rcolor, label=names[i],pos=position[i])

    # 添加边并设置边属性
    for i in range(len(matrix)):
        for j in range(i + 1, len(matrix)):
            if matrix[i][j] != 0 and matrix[i][j]!=INF:
                if i==53 and j==63:
                    rcolor = "#6a2947"
                elif 1 in route[i] and 1 in route[j]:
                    rcolor = "#ecc647"
                elif 2 in route[i] and 2 in route[j]:
                    rcolor = "#365c88"
                elif 3 in route[i] and 3 in route[j]:
                    rcolor = "#e78e49"
                elif 4 in route[i] and 4 in route[j]:
                    rcolor = "#32723b"
                elif 5 in route[i] and 5 in route[j]:
                    rcolor = "#bb3140"
                elif 6 in route[i] and 6 in route[j]:
                    rcolor = "#6a2947"
                G.add_edge(i, j, color=rcolor)

    # 设置画布大小，定义节点的位置
    #fig, ax = plt.subplots(figsize=(16,8))
    pos = nx.get_node_attributes(G,'pos')

    # 绘制图形
    node_colors = [G.nodes[n]["color"] for n in G.nodes()]
    edge_colors = [G.edges[e]["color"] for e in G.edges()]

    # 绘制节点和边，并调整节点大小
    nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=100)
    nx.draw_networkx_edges(G, pos, edge_color=edge_colors,arrows=False,width=2)

    # 添加节点标签
    labels = nx.get_node_attributes(G, 'label')
    nx.draw_networkx_labels(G, pos, labels, font_size=10, font_color="black",font_family='SimHei',font_weight='bold')

    plt.axis("off")  # 隐藏坐标轴

    # 调整绘图区域与画布边界之间的距离
    plt.subplots_adjust(left=0.0001, right=0.9999, top=0.9999, bottom=0.0001)
    #plt.savefig("graph.png")  # 保存图形为 png 格式的图片
    #plt.show()  # 显示图形
    fig.canvas.draw()

def floyd(matrix,weight):
    # 用Floyd算法计算每两个节点间的最短路径。weight==0时权重为路程，weight==1时权重为时间,weight==3时权重为1
    n = len(matrix)

    # 初始化距离矩阵和路径矩阵
    dist = []
    path = []
    for i in range(n):
        dist_row = []  # 存储节点i到各个节点的最短距离
        path_row = []  # 存储节点i到各个节点的最短路径上的下一个节点
        for j in range(n):
            if isinstance(matrix[i][j], list):
                dist_row.append(matrix[i][j][weight])  # 邻接矩阵中为列表时，用对应的权重
                path_row.append(j)  # 有连接时，下一个节点为终点j
            else:
                dist_row.append(matrix[i][j])
                path_row.append(-1)  #无连接时，设置为-1
        dist.append(dist_row)
        path.append(path_row)

    # 使用Floyd算法计算最短路径
    for k in range(n):
        for i in range(n):
            for j in range(n):
                if dist[i][j] > dist[i][k] + dist[k][j]:
                    dist[i][j] = dist[i][k] + dist[k][j]
                    path[i][j] = path[i][k]
    return dist, path

def showGraphPath(p,fig): #构建地铁线路的路径图
    # 创建有向图
    G = nx.DiGraph()
    # 添加节点并设置节点属性
    for i in p:
        if route[i][0] == 1:
            rcolor = "#f9eec8"
        elif route[i][0] == 2:
            rcolor = "#d7dee7"
        elif route[i][0] == 3:
            rcolor = "#fae8db"
        elif route[i][0] == 4:
            rcolor = "#d6e3d8"
        elif route[i][0] == 5:
            rcolor = "#f8eaec"
        elif route[i][0] == 6:
            rcolor = "#f0eaed"
        G.add_node(i, color=rcolor, label=names[i],pos=position[i])

    # 添加边并设置边属性
    pathroute = []  # 记录每条边的地铁路线
    for i in range(len(p)-1):
        if (p[i]==53 and p[i+1]==63) or (p[i]==63 and p[i+1]==53):  # 燕塘到天河客运站是特殊情况
            pathroute.append(6)
            rcolor = "#6a2947"
        elif 1 in route[p[i]] and 1 in route[p[i+1]]:
            pathroute.append(1)
            rcolor = "#ecc647"
        elif 2 in route[p[i]] and 2 in route[p[i+1]]:
            pathroute.append(2)
            rcolor = "#365c88"
        elif 3 in route[p[i]] and 3 in route[p[i+1]]:
            pathroute.append(3)
            rcolor = "#e78e49"
        elif 4 in route[p[i]] and 4 in route[p[i+1]]:
            pathroute.append(4)
            rcolor = "#32723b"
        elif 5 in route[p[i]] and 5 in route[p[i+1]]:
            pathroute.append(5)
            rcolor = "#bb3140"
        elif 6 in route[p[i]] and 6 in route[p[i+1]]:
            pathroute.append(6)
            rcolor = "#6a2947"
        G.add_edge(p[i], p[i+1], color=rcolor)

    # 设置画布大小，定义节点的位置
    #fig, ax = plt.subplots(figsize=(16,8))
    pos = nx.get_node_attributes(G,'pos')

    # 绘制图形
    node_colors = [G.nodes[n]["color"] for n in G.nodes()]
    edge_colors = [G.edges[e]["color"] for e in G.edges()]

    # 绘制节点和边，并调整节点大小
    nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=100)
    nx.draw_networkx_edges(G, pos, edge_color=edge_colors,arrows=False,width=2)

    # 添加节点标签
    labels = nx.get_node_attributes(G, 'label')
    nx.draw_networkx_labels(G, pos, labels, font_size=13, font_color="black",font_family='SimHei',font_weight='bold')

    plt.axis("off")  # 隐藏坐标轴

    # 调整绘图区域与画布边界之间的距离
    plt.subplots_adjust(left=0.0001, right=0.9999, top=0.9999, bottom=0.0001)
    #plt.savefig("path.png")  # 保存图形为 png 格式的图片
    #plt.show()  # 显示图形
    fig.canvas.draw()
    return pathroute

def UIdesign():
    def on_submit():
        selected_option = var.get()
        if selected_option == "distance":
            weight = 0
        elif selected_option == "time":
            weight = 1
        elif selected_option == "station":
            weight = 2
        dist, path = floyd(matrix, weight)  # 调用Floyd函数计算最短路径
        notice_text.set("")
        start_point = start_entry.get()
        end_point = end_entry.get()
        if start_point == "":
            notice_text.set("起点站点不能为空")
        elif end_point =="":
            notice_text.set("终点站点不能为空")
        elif start_point not in names:
            notice_text.set("没有此起点站点！")
        elif end_point not in names:
            notice_text.set("没有此终点站点!")
        elif start_point==end_point:
            notice_text.set("起点与终点相同!")
        else:
            start_index = names.index(start_point)
            end_index = names.index(end_point)

            if dist[start_index][end_index] != INF:
                #print(f"节点{start_index}到节点{end_index}的最短距离为{dist[start_index][end_index]}，路径为：", end="")
                p = [start_index]  # 存储路径index的列表
                k = path[start_index][end_index]  # 路径上的下一个节点
                while k != end_index:  # 回溯路径直到到达终点j
                    p.append(k)
                    k = path[k][end_index]
                p.append(end_index)  # 将终点j添加到路径中
                pname=[]  # 存储路径名称
                for each in p:
                    pname.append(names[each])
                #print(pname)

                ax.clear()

                pathroute = showGraphPath(p,fig)
                #print(pathroute)
                #print(p)

                if weight ==0:
                    shortest_path = dist[start_index][end_index]
                elif weight == 1:
                    shortest_time = dist[start_index][end_index]
                elif weight == 2:
                    fewest_station = dist[start_index][end_index]

                pathtext = ""
                for i in range(len(pname)):
                    if i==0:
                        pathtext += "(" + str(pathroute[i]) + "号线)\n" + pname[0] + "-->\n"
                    elif i==len(pname)-1:
                        pathtext += pname[i] + "\n"
                    elif i==len(pname)-2:
                        pathtext += pname[i] + "-->\n"
                    elif pathroute[i] != pathroute[i+1]:
                        pathtext += pname[i] + "-->\n"
                        pathtext += "(从" + str(pathroute[i]) + "号线换乘到" + str(pathroute[i + 1]) + "号线)\n"
                        if weight==1:
                            shortest_time +=2  # 默认换乘要用2分钟
                    else:
                        pathtext += pname[i] + "-->\n"

                if weight==0:
                    pathtext += "\n最短路程为：{:.2f}".format(shortest_path) + "km"
                if weight==1:
                    pathtext += "\n最短时间为：" + str(shortest_time) + "分钟"
                if weight ==2:
                    pathtext += "\n最少站数为：" + str(fewest_station) + "站"

                output_path.set(pathtext)


    def on_reset():
        start_entry.delete(0,tk.END)
        end_entry.delete(0,tk.END)
        ax.clear()
        showGraphTotal(matrix, fig)
        notice_text.set("")
        output_path.set("")

    # 创建主窗口
    root = tk.Tk()
    root.title("地铁路径规划")
    root.geometry("2400x1400")
    root.configure(bg='white')

    # 创建起点和终点的输入框和标签
    start_label = tk.Label(root, text="起点:",font=('SimHei',18),background='white')
    start_label.place(x=50,y=50)
    start_entry = tk.Entry(root,width=15,font=('SimHei',17))
    start_entry.place(x=120,y=50)

    end_label = tk.Label(root, text="终点:",font=('SimHei',18),background='white')
    end_label.place(x=50,y=110)
    end_entry = tk.Entry(root,width=15,font=('SimHei',17))
    end_entry.place(x=120,y=110)

    # 对错误输入的提示语输出文字
    notice_text = tk.StringVar()
    notice_text.set("")
    notice_label = tk.Label(root, textvariable=notice_text, font=("SimHei", 18), fg='red', background='white')
    notice_label.place(x=80, y=350)

    # 输出路线
    output_path = tk.StringVar()
    output_path.set("")
    output_label = tk.Label(root, textvariable=output_path, font=("SimHei", 16), background='white')
    output_label.place(x=50, y=370)

    # 创建变量来存储选择的选项
    var = tk.StringVar(value="distance")
    # 创建选择框
    option1 = tk.Radiobutton(root, text="最短路程", variable=var, value="distance",width=10,height=2,font=('SimHei',16),background='white')
    option1.place(x=40,y=160)
    option2 = tk.Radiobutton(root, text="最短时间", variable=var, value="time",width=10,height=2,font=('SimHei',16),background='white')
    option2.place(x=170,y=160)
    option3 = tk.Radiobutton(root, text="最少站数", variable=var, value="station",width=10,height=2,font=('SimHei',16),background='white')
    option3.place(x=120,y=210)

    # 创建查询按钮
    submit_button = tk.Button(root, text="查询", command=on_submit,width=10,height=2,font=('SimHei',15))
    submit_button.place(x=50,y=280)

    #创建重置按钮
    reset_button = tk.Button(root,text="重置",command=on_reset,width=10,height=2,font=('SimHei',15))
    reset_button.place(x=200,y=280)


    # 插入图片
    image = Image.open("symbols.png")
    resized_image = image.resize((1200,45))
    photo = ImageTk.PhotoImage(resized_image)
    img_label = tk.Label(root, image=photo,background='white')
    img_label.image = photo  # 保持对图片对象的引用，避免被垃圾回收
    img_label.place(x=800,y=1230)

    fig, ax = plt.subplots(figsize=(20, 12))
    # 创建一个canvas来容纳matplotlib图形
    canvas = FigureCanvasTkAgg(fig, master=root)
    showGraphTotal(matrix,fig)
    canvas.get_tk_widget().place(x=350,y=20)

    # 运行主循环
    root.mainloop()


names, route, adj, distance,time,position = get_data()

matrix = create_adj_matrix(names)
n = len(matrix)

UIdesign()