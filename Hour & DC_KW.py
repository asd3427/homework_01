import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from sklearn.linear_model import LinearRegression
from sklearn .model_selection import train_test_split
from sklearn.preprocessing import PolynomialFeatures
from openpyxl import Workbook
from sklearn.metrics import mean_squared_error
import xlwt
#上面 导入了我们需要用到的 库(api)

Test_Data =  pd.read_excel('HourlyWP .xlsx','testing dataset')
Trin_Data = pd.read_excel('HourlyWP .xlsx','training dataset') # 读取 资料
Test_Data.head() #显示资料的 内容 (标题栏)
Test_Data.info()# 显示资料的详细资讯



x_hour = Trin_Data[['hour']] 
y_hour= Trin_Data[['DC_KW']]
X_test_hour =Test_Data[['hour']] #自变量
y_test_hour =Test_Data[['DC_KW']] 
plt.plot(X_test_hour,y_test_hour , 'r.')#'C.'  点点的颜色
 # 上面试做资料的分类 
model = LinearRegression()#设定MODEL为线性回归
model.fit(x_hour, y_hour)
X2_hour = X_test_hour # 自变量
y2_hour = model.predict(X2_hour)
#画出图形
plt.figure()
plt.grid(True) #开启格线
plt.title('liner')
plt.xlabel('hour') 
plt.ylabel('DC_KW') 
plt.plot(X2_hour, y2_hour, 'b-',linewidth=3)# 线的颜色
plt.plot(X_test_hour,y_test_hour , 'r.')#'C.'  颜色
plt.show() 


#xx = np.linspace(0, 1200)   # 均分命令 将0~1200 均分成若干份  m这里默认是50

quadratic_featurizer_hour = PolynomialFeatures(degree=2)  #实例化一个二次多项式特征实例
X_train_quadratic_hour = quadratic_featurizer_hour.fit_transform(x_hour) #用二次多项式对样本X值做变换
X_test_quadratic_hour = quadratic_featurizer_hour.transform(X_test_hour)  #把训练好X值的多项式特征实例应用到一系列点上,形成矩阵
regressor_quadratic = LinearRegression()  # 创建一个线性回归实例
regressor_quadratic.fit(X_train_quadratic_hour, y_hour) # 以多项式变换后的x值为输入，代入线性回归模型做训练
prediction_hour = regressor_quadratic.predict(X_test_quadratic_hour) #获得预测结果

xx_hour = np.linspace(6, 18,100)
test_interp_quad = quadratic_featurizer_hour.transform(xx_hour.reshape(-1,1))
prediction_hour_interp = regressor_quadratic.predict(test_interp_quad)
plt.figure()
plt.grid(True) #开启格线
plt.title('poly_2')
plt.xlabel('hour') 
plt.ylabel('DC_KW') 
plt.plot(X2_hour, y2_hour, 'b-',linewidth=3)# 线的颜色
plt.plot(xx_hour, prediction_hour_interp, 'g-',linewidth=3)# 线的颜色
plt.plot(X_test_hour,y_test_hour , 'r.')#'C.'   颜色
plt.show() 

score_one = model.score(X_test_hour, y_test_hour)
score_two = regressor_quadratic.score(X_test_quadratic_hour, y_test_hour)
print( '一次線性回歸     r-squared',model.score(X_test_hour, y_test_hour))
X_test_quadratic_hour = quadratic_featurizer_hour.transform(X_test_hour)
print ('二次線性回歸     r-squared', regressor_quadratic.score(X_test_quadratic_hour, y_test_hour))
loss_hour= mean_squared_error(y_test_hour, prediction_hour)

print('loss of hour is :',loss_hour)

#画图
plt.subplot(211)
plt.grid(True) #开启格线
plt.title('point')
plt.xlabel('Predicted')
plt.ylabel('Real')
plt.plot(prediction_hour,y_test_hour, 'c.', lw=2)


plt.subplot(212)
plt.grid(True) #开启格线
plt.title('liner')
plt.xlabel('Predicted')
plt.ylabel('Real')
plt.plot([y_test_hour.min(), y_test_hour.max()], [y_test_hour.min(), y_test_hour.max()], 'r--', lw=2)
plt.tight_layout()
plt.show()

testdataX = np.array(X_test_hour)    #自变数
testdataY = np.array(y_test_hour)    #一变数
Prediction_all_hour = np.array(prediction_hour) #预测结果



## 印出资料
book = xlwt.Workbook()
sheet = book.add_sheet('HW',cell_overwrite_ok=True)
sheet.write(0,0,'hour 自变数') #自變數
sheet.write(0,1,'Real DC_KW 因变数') #因變數
sheet.write(0,2,'Predicted DC_KW 预测结果') #預測結果
sheet.write(0,3,'loss 误差值') #預測誤差值之平方和
sheet.write(0,4,'一次回归') #績效
sheet.write(0,5,'二次回归')
for i in range(0,1014):
   sheet.write(i+1,0,float(testdataX[i][0]))
   sheet.write(i+1,1,testdataY[i][0])
   sheet.write(i+1,2,Prediction_all_hour[i][0])
   

   
sheet.write(1,3,loss_hour)
sheet.write(1,4,score_one)
sheet.write(1,5,score_two)


book.save('linear_B.xls')


#
