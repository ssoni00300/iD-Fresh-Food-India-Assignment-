import pandas as pd
import matplotlib.pyplot as plt


##Task 1
File_Path = 'Superstore.xls'
df = pd.read_excel(File_Path)
print(df.head())


##Task 2
Orders_df = pd.read_excel('Superstore.xls', sheet_name='Orders')
People_df = pd.read_excel('Superstore.xls', sheet_name='People')
Returns_df= pd.read_excel('Superstore.xls', sheet_name='Returns')

##Task 3
Merged_df = pd.merge(Orders_df, People_df, on='Region')
Region_Sales = Merged_df.groupby('Region')['Sales'].sum()
print(Region_Sales)


##Task 4
Product_Sales = Orders_df.groupby('Product Name')['Sales'].sum()
Product_Quantity = Orders_df.groupby('Product Name')['Quantity'].sum()

# Plot the sales and quantity in a bar chart
plt.figure(100,figsize=(15, 8))
Product_Sales.plot(kind='bar', color='Yellow', alpha=0.7, label='Sales')
Product_Quantity.plot(kind='bar', color='Red', alpha=0.7, label='Quantity')
plt.xlabel('Product Name')
plt.ylabel('Value')
plt.title('Product Name Wise Sales and Quantity')
plt.legend()
plt.show()

# Displaying the sales and quantity in a table format
Product_Data = pd.DataFrame({'Sales': Product_Sales, 'Quantity': Product_Quantity})
print(Product_Data)


##Task 5
Product_Sales_quantity = Orders_df.groupby('Product Name')['Quantity'].sum()
Top_10_Products = Product_Sales_quantity.sort_values(ascending=False).head(10)
print(Top_10_Products)


##Task 6
Category_Sales = Orders_df.groupby('Category')['Sales'].sum()
print(Category_Sales)

##Task 7
Product_Category_Sales = Orders_df.groupby(['Product', 'Category'])['Sales'].sum()
print(Product_Category_Sales)

##Task 8
Customer_Sales = Orders_df.groupby(['Customer ID', 'Customer Name'])['Sales'].sum()
Top_10_Customers = Customer_Sales.sort_values(ascending=False).head(10)
print(Top_10_Customers)


##Task 9
Merged_df = pd.merge(Orders_df, Returns_df, on='Order ID', how='left')
Customer_Return_Count = Merged_df.groupby(['Customer ID', 'Customer Name'])['Returned'].count()
Sorted_Customers = Customer_Return_Count.sort_values(ascending=False)
print(Sorted_Customers)


##PLOTS

plt.figure(200,figsize=(15, 8))
Category_Sales.plot(kind='bar', color='Yellow', alpha=0.7, label='Sales')
plt.xlabel('Category Name')
plt.ylabel('Value')
plt.title('Category Name Wise Sales')
plt.legend()
plt.show()
