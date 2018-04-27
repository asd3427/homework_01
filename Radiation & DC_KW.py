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




x = Trin_Data[['radiation']] 
y= Trin_Data[['DC_KW']]
X_test =Test_Data[['radiation']] #自变量
y_test =Test_Data[['DC_KW']] 
plt.plot(X_test,y_test , 'r.')#'C.'  点点的颜色
 # 上面试做资料的分类 
model = LinearRegression()#设定MODEL为线性回归
model.fit(x, y)
X2 = X_test # 自变量
y2 = model.predict(X2)
#画出图形
plt.subplot(211)
plt.grid(True) #开启格线
plt.title('liner')
plt.xlabel('radiation') 
plt.ylabel('DC_KW') 
plt.plot(X2, y2, 'b-',linewidth=3)# 线的颜色
plt.subplot(212)
plt.grid(True) #开启格线
plt.title('point')
plt.xlabel('radiation') 
plt.ylabel('DC_KW') 
plt.plot(X_test,y_test , 'r.')#'C.'  颜色
plt.tight_layout()

plt.show() 

#xx = np.linspace(0, 1200)   # 均分命令 将0~1200 均分成若干份  m这里默认是50

quadratic_featurizer = PolynomialFeatures(degree=2)  #实例化一个二次多项式特征实例
X_train_quadratic = quadratic_featurizer.fit_transform(x) #用二次多项式对样本X值做变换
X_test_quadratic = quadratic_featurizer.transform(X_test)  #把训练好X值的多项式特征实例应用到一系列点上,形成矩阵
regressor_quadratic = LinearRegression()  # 创建一个线性回归实例
regressor_quadratic.fit(X_train_quadratic, y) # 以多项式变换后的x值为输入，代入线性回归模型做训练
prediction = regressor_quadratic.predict(X_test_quadratic) #获得预测结果

xx = np.linspace(0, 1013,100)
test_interp_quad = quadratic_featurizer.transform(xx.reshape(-1,1))
prediction_interp = regressor_quadratic.predict(test_interp_quad)
plt.figure()
plt.grid(True) #开启格线
plt.title('poly_2')
plt.xlabel('radiation') 
plt.ylabel('DC_KW') 
plt.plot(X2, y2, 'b-',linewidth=3)# 线的颜色
plt.plot(xx, prediction_interp, 'y-',linewidth=3)# 线的颜色

plt.show() 










score_one = model.score(X_test, y_test)
score_two = regressor_quadratic.score(X_test_quadratic, y_test)
print( '一次線性回歸     r-squared',model.score(X_test, y_test))
X_test_quadratic = quadratic_featurizer.transform(X_test)
print ('二次線性回歸     r-squared', regressor_quadratic.score(X_test_quadratic, y_test))
loss= mean_squared_error(y_test, prediction)

print('loss is : ',loss)

plt.subplot(211)
plt.grid(True) #开启格线
plt.title('point')
plt.xlabel('Predicted')
plt.ylabel('Real')
plt.plot(y_test,prediction, 'c.', lw=2)

plt.subplot(212)
plt.grid(True) #开启格线
plt.title('liner')
plt.xlabel('Predicted')
plt.ylabel('Real')
plt.plot([y_test.min(), y_test.max()], [y_test.min(), y_test.max()], 'k-', lw=2)
plt.tight_layout()
plt.show()

#sns.distplot(X_test) #看實際值及預測值之間的殘差分佈圖
testdataX = np.array(X_test)    #自变数
testdataY = np.array(y_test)    #一变数
Prediction_all = np.array(prediction) #预测结果
    
book = xlwt.Workbook()
sheet = book.add_sheet('HW',cell_overwrite_ok=True)
sheet.write(0,0,'Radiation 自变数') #自變數
sheet.write(0,1,'Real DC_KW 因变数') #因變數
sheet.write(0,2,'Predicted DC_KW 预测结果') #預測結果
sheet.write(0,3,'loss 误差 ') #預測誤差值之平方和
sheet.write(0,4,'Score_one 一次回归') #績效
sheet.write(0,5,'Score_two 二次回归 ' ) #績效
for i in range(0,1013):
    sheet.write(i+1,0,testdataX[i][0])
    sheet.write(i+1,1,testdataY[i][0])
    sheet.write(i+1,2,Prediction_all[i][0])
   
    
sheet.write(1,3,loss)
sheet.write(1,4,score_one)
sheet.write(1,5,score_two)


book.save('linear_A.xls')




