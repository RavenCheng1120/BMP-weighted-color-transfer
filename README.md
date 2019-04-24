# 照片色彩轉換，加上權重  
Pictures color transfer with weight.    

#### 讀取照片的RGB值，平均值，標準差    
```python
sImg = cv2.imread('./source/'+source)
sImg = cv2.cvtColor(sImg, cv2.COLOR_BGR2LAB)
sMean, sStd = cv2.meanStdDev(sImg)
sMean = np.hstack(np.around(sMean, decimals=2))
sStd = np.hstack(np.around(sStd, decimals=2))
```

#### 色彩轉換的公式    
```python
#一般轉換公式
R = (S - sMean) * (tStd / sStd) + tMean
#加上權重w
s = ((w*trStd[k]+(1-w)*sStd[k]) / sStd[k]) * (s-sMean[k]) + trMean[k]*w + (1-w)*sMean[k]
```

#### 計算COLOR DISTANCE的公式    
```python
if k==0:
  ESI+=3*(s-sImg[i,j,k])*(s-sImg[i,j,k])
  ETI+=3*(s-trImg[i,j,k])*(s-trImg[i,j,k])
elif k==1:
  ESI+=4*(s-sImg[i,j,k])*(s-sImg[i,j,k])
  ETI+=4*(s-trImg[i,j,k])*(s-trImg[i,j,k])
elif k==2:
  ESI+=2*(s-sImg[i,j,k])*(s-sImg[i,j,k])
  ETI+=2*(s-trImg[i,j,k])*(s-trImg[i,j,k])
ESI = ESI/(height*width)
ETI = ETI/(height*width)
TCD = abs(ESI - ETI)
```    
         
weight從0.00~1.00，輸出一百零一張color transfer結果到excel中，包含mean,std,color distance,TCD和optimal的解。        
使用pyinstaller將python打包成exe檔 `pyinstaller -F .\4105056005-DCSA-07.py`     
excel使用`from openpyxl import Workbook,load_workbook`        
    
圖源:https://unsplash.com/
