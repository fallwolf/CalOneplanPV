import numpy as np
import xlrd
import math
import tkinter as tk
from tkinter import filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def caldata(file):
  # 调试时显示当前计算的文件
  #print(file)
  workbook = xlrd.open_workbook(file)

  # 选择第一个工作表
  worksheet = workbook.sheet_by_index(0)
  datalen=len(worksheet.col_values(2))

  # 读取第2列数据--x
  x = worksheet.col_values(2)[3:datalen]
  ax = np.array(x)
  # 读取第3列数据--y
  y = worksheet.col_values(3)[3:datalen]
  ay = np.array(y)
  # 读取第3列数据--z
  z = worksheet.col_values(4)[3:datalen]
  az = np.array(z)
  N = datalen-3
  # ------------------------构建系数矩阵-----------------------------
  A = np.array([[sum(ax ** 2), sum(ax * ay), sum(ax)],
              [sum(ax * ay), sum(ay ** 2), sum(ay)],
              [sum(ax), sum(ay), N]])

  B = np.array([[sum(ax * az), sum(ay * az), sum(az)]])

  # 求解
  X = np.linalg.solve(A, B.T)

  # 计算点到平面距离
  pvbase=np.ones(len(ax))
  i=0
  while i<len(ax):
      pvbase[i]=(X[0]*ax[i]+X[1]*ay[i]-az[i]+X[2])/math.sqrt(X[0]*X[0]+X[1]*X[1]+1)
      i+=1

  PV =max(pvbase)-min(pvbase)
  return PV,x,y,z,ax,ay,X

#窗口运行代码


def open_file():
    filetypes = (
        ('97 Excel Files', '*.xls'),
        ('CSV Files', '*.csv'),
        ('Excel Files', '*xlsx')
    )

    file_path = filedialog.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes
    )

    if file_path:
        with open(file_path, 'r') as file:
            #print(file.name)
            dataR=caldata(file.name)#PV,x,y,z,ax,ay,X
            text.insert('1.0', '平面度：'+str("{:.1f}".format(dataR[0]*1000))+' um')
            canvas = tk.Canvas(root, width=500, height=500)
            canvas.pack()
            # 创建matplotlib的Figure和Axes对象
            fig1 = plt.figure()
            ax1 = fig1.add_subplot(111, projection='3d')
            ax1.set_xlabel("x")
            ax1.set_ylabel("y")
            ax1.set_zlabel("z")
            ax1.scatter(dataR[1], dataR[2], dataR[3], c='r', marker='o')
            x_p = np.linspace(min(dataR[4]), max(dataR[4]), 100)
            y_p = np.linspace(min(dataR[5]), max(dataR[5]), 100)
            x_p, y_p = np.meshgrid(x_p, y_p)
            z_p = dataR[6][0] * x_p + dataR[6][1] * y_p + dataR[6][2]
            ax1.plot_wireframe(x_p, y_p, z_p, rstride=10, cstride=10)
            plt.savefig(file.name + '.jpg', dpi=600)

            # 将matplotlib的画布嵌入到tkinter口中
            canvas = FigureCanvasTkAgg(fig1, master=canvas)
            canvas.draw()
            canvas.get_tk_widget().pack()

def quit_app():
    root.quit()

root = tk.Tk()
root.title('计算平面度程序@20230819')

text = tk.Text(root, height=2)
text.pack()

button = tk.Button(root, text='打开文件', command=open_file)
button.pack()

exit_button = tk.Button(root, text="退出", command=quit_app)
exit_button.pack()

root.mainloop()
#plt.show()
