from sklearn.linear_model import LinearRegression,Lasso,Ridge,ElasticNet,LogisticRegression
from sklearn.model_selection import cross_val_score,train_test_split,GridSearchCV,RandomizedSearchCV
from sklearn.metrics import confusion_matrix,mean_squared_error,mean_absolute_error,mean_absolute_percentage_error,r2_score,accuracy_score,f1_score,precision_score,recall_score
from sklearn.tree import DecisionTreeClassifier,export_graphviz
from sklearn import tree
from sklearn.preprocessing import MinMaxScaler,StandardScaler,LabelEncoder
from sklearn.ensemble import RandomForestClassifier,RandomForestRegressor   
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import xgboost as xgb
import os
import pickle

#class created
class PredictiveMaintenance():
    def __init__(self):
        os.chdir('D:\PythonProjeleri\DataScienceMachineLearning\Docs')
        file = "predictive_maintenance.csv"
        try:
            data = pd.read_csv(file)
        except FileNotFoundError:
            print("File not found.")
            return
        df=pd.DataFrame(data)
        #preprocessing operations for graphics
        df.columns=['UDI', 'Product ID', 'Type', 'Air temperature [°C]',
       'Process temperature [°C]', 'Rotational speed [rpm]', 'Torque [Nm]',
       'Tool wear [min]', 'Failure', 'Failure Type']
        df.drop(columns=['Product ID','UDI','Type'],inplace=True)
        #print(pd.crosstab(df["Failure"], df["Failure Type"]))
        df.drop(df.loc[(df["Failure"] == 1) & (df["Failure Type"] == "No Failure")].index, inplace=True)
        df['Air temperature [°C]']=df['Air temperature [°C]']-273.15
        df['Process temperature [°C]']=df['Process temperature [°C]']-273.15
        print(df.head())
        self.dfFailBin=df.drop('Failure Type',axis=1)
        self.dfFailMulti=df.drop('Failure',axis=1)
    def correlationGraphic(self):
        #making correlation graphic
        data=self.dfFailBin.corr(method='pearson')
        sns.heatmap(data=data,annot=True)
        plt.show()
    def histogramGraphic(self):
        #making histogram graphic
        fig, axes = plt.subplots(2, 3, figsize=(10, 10)) 
        features = self.dfFailBin.columns
        axes = axes.flatten()
        for i, feature in enumerate(features):
            sns.histplot(self.dfFailBin[feature], kde=True, ax=axes[i])
            axes[i].set_title(f'{feature}')
        plt.tight_layout()
        plt.show()
    def boxplotGraphic(self):
        #making boxplot graphic
        fig, axes = plt.subplots(2, 3, figsize=(10, 10)) 
        features = self.dfFailBin.columns
        axes = axes.flatten()
        for i, feature in enumerate(features):
            sns.boxplot(self.dfFailBin[feature], ax=axes[i])
            axes[i].set_title(f'{feature}')
        plt.tight_layout()
        plt.show()
    def logisticRegression(self):
        lr=LogisticRegression()
        mm=MinMaxScaler()
        y=self.dfFailBin['Failure']
        x=self.dfFailBin.drop('Failure',axis=1)
        xScaled=mm.fit_transform(x)
        xScaledDf=pd.DataFrame(xScaled,columns=x.columns)
        x_train,x_test,y_train,y_test=train_test_split(xScaledDf,y,test_size=0.2,random_state=42)
        model=lr.fit(x_train,y_train)
        y_pred=model.predict(x_test)
        print("model score: ",model.score(x_test,y_test))
        print("accuracy: ",accuracy_score(y_test,y_pred))
        print("precision score: ", precision_score(y_test,y_pred))
        print("recall score: ",recall_score(y_test,y_pred))
        print("f1 score: ",f1_score(y_test,y_pred))
    def decisionTree(self):
        dt=DecisionTreeClassifier()
        mm=MinMaxScaler()
        y=self.dfFailBin['Failure']
        x=self.dfFailBin.drop('Failure',axis=1)
        xScaled=mm.fit_transform(x)
        xScaledDf=pd.DataFrame(xScaled,columns=x.columns)
        x_train,x_test,y_train,y_test=train_test_split(xScaledDf,y,test_size=0.2,random_state=42)
        model=dt.fit(x_train,y_train)
        y_pred=model.predict(x_test)
        print("model score: ",model.score(x_test,y_test))
        print("accuracy: ",accuracy_score(y_test,y_pred))
        print("precision score: ", precision_score(y_test,y_pred))
        print("recall score: ",recall_score(y_test,y_pred))
        print("f1 score: ",f1_score(y_test,y_pred))
    def boostingClassifier(self):
        boost=xgb.XGBClassifier()
        mm=MinMaxScaler()
        self.dfFailBin.columns=['Air temperature', 'Process temperature',
       'Rotational speed', 'Torque', 'Tool wear', 'Failure']
        y=self.dfFailBin['Failure']
        x=self.dfFailBin.drop('Failure',axis=1)
        xScaled=mm.fit_transform(x)
        xScaledDf=pd.DataFrame(xScaled,columns=x.columns)
        x_train,x_test,y_train,y_test=train_test_split(xScaledDf,y,test_size=0.2,random_state=42)
        model=boost.fit(x_train,y_train)
        y_pred=model.predict(x_test)
        print("model score: ",model.score(x_test,y_test))
        print("accuracy: ",accuracy_score(y_test,y_pred))
        print("precision score: ", precision_score(y_test,y_pred))
        print("recall score: ",recall_score(y_test,y_pred))
        print("f1 score: ",f1_score(y_test,y_pred))
        
    def randomForestClassifier(self):
        rfc=RandomForestClassifier()
        mm=MinMaxScaler()
        y=self.dfFailBin['Failure']
        x=self.dfFailBin.drop('Failure',axis=1)
        #Applying minmaxscaler
        xScaled=mm.fit_transform(x)
        xScaledDf=pd.DataFrame(xScaled,columns=x.columns)
        x_train,x_test,y_train,y_test=train_test_split(xScaledDf,y,test_size=0.3,random_state=22)
        model=rfc.fit(x_train,y_train)
        y_pred=model.predict(x_test)
        #Model scores
        print("model score: ",model.score(x_test,y_test))
        print("accuracy: ",accuracy_score(y_test,y_pred))
        print("precision score: ", precision_score(y_test,y_pred))
        print("recall score: ",recall_score(y_test,y_pred))
        print("f1 score: ",f1_score(y_test,y_pred)) 
        #Randomgridsearch CV and model tuning
        rand_params = { "n_estimators": np.arange(10, 1000, 50),
                       "max_depth": [None, 5, 10], 
                       "min_samples_split": np.arange(2, 20, 4), 
                       "min_samples_leaf": np.arange(1, 10, 2) }
        rs_rf = RandomizedSearchCV(rfc,param_distributions=rand_params,cv=5,verbose=True,
        return_train_score=True,n_iter=20,scoring='accuracy')
        model_tuned=rs_rf.fit(x_train,y_train)
        print('Best hyper parameter:', rs_rf.best_params_)
        print("Score after tuning:", model_tuned.score(x_test,y_test))
        #Confusion matrix 
        print(confusion_matrix(y_test,y_pred))
        sns.set_theme(font_scale=1.5)
        fig, ax = plt.subplots(figsize=(3,3))
        ax = sns.heatmap(confusion_matrix(y_test, y_pred),
                     annot=True,
                     cbar=False)
        plt.xlabel("Predicted Values")
        plt.ylabel("True Values")
        plt.show()
        #Tuned model's cross-validation scores
        rfcTuned=RandomForestClassifier(n_estimators=660,min_samples_split=2,
                                        min_samples_leaf=1,max_depth=None)
        cv_acc=cross_val_score(rfcTuned,x,y,cv=5,scoring='accuracy')
        print("cross-validation accuracy: ",np.mean(cv_acc))
        cv_recall=cross_val_score(rfcTuned,x,y,cv=5,scoring='recall')
        print("cross-validation recall: ",np.mean(cv_recall))
        cv_f1=cross_val_score(rfcTuned,x,y,cv=5,scoring='f1')
        print("cross-validation f1: ",np.mean(cv_f1))
        cv_prec=cross_val_score(rfcTuned,x,y,cv=5,scoring='precision')
        print("cross-validation precision: ",np.mean(cv_prec))
        #Getting feature importances for tuned model
        rfc=RandomForestClassifier()
        rfc.fit(x_train,y_train)
        print("feature importances: ",rfc.feature_importances_) 
        features=dict(zip(x.columns,rfc.feature_importances_))
        print("features: ",features)
        feature_df = pd.DataFrame(features, index=[0])
        feature_df.T.plot(kind="bar",
                  title="Feature Importance",
                  legend=False)
        plt.show()
        #Saving model
        with open('random_forest_model.pkl', 'wb') as file: 
             pickle.dump(rfcTuned, file)


predictiveMaintenance=PredictiveMaintenance()  

while True:
    print("Select an option:")
    print("1. Correlation graph, 2. Histogram graph, 3. Boxplot graph, 4. Logistic regression scores")
    print("5. Decision tree scores, 6. Boosting classifier scores, 7. Random forest classifier scores")
    print("8-Exit")
    userChoice = input("Enter your choice: ")
    if userChoice == "1":
            print("Correlation graph selected")
            predictiveMaintenance.correlationGraphic()
    elif userChoice == "2":
            print("Histogram graph selected")
            predictiveMaintenance.histogramGraphic()
    elif userChoice == "3":
            print("Boxplot graph selected")
            predictiveMaintenance.boxplotGraphic()
    elif userChoice == "4":
            print("Logistic regression scores selected")
            predictiveMaintenance.logisticRegression()
    elif userChoice == "5":
            print("Decision tree scores selected")
            predictiveMaintenance.decisionTree()
    elif userChoice == "6":
            print("Boosting classifier scores selected")
            predictiveMaintenance.boostingClassifier()
    elif userChoice == "7":
            print("Random forest classifier scores selected")
            predictiveMaintenance.randomForestClassifier()
    else:
            print("App closed.")
            break
    