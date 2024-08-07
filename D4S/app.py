# %%
#importation de différents modules
import numpy as np 
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# %%
#chargement des données customers 
customers = pd.read_excel("sales.xlsx", "Customers", header=None, skiprows=1, 
                          names=['customerID', 'customerName', 'size', 'capital'])
missingVals = ['na', 'NaN', ' ', None, np.nan, 'azerty', '1111']
mask_customers = customers.isin(missingVals)
customers.mask(mask_customers, None, inplace=True)
customers.loc[(customers['size'].isna()) & (customers['capital'] < 10), 'size'] = 'Micro'
customers.loc[(customers['size'].isna()) & (customers['capital'] < 100), 'size'] = 'Small'
customers.loc[(customers['size'].isna()) & (customers['capital'] < 1000), 'size'] = 'Medium'
customers.loc[(customers['size'].isna()), 'size'] = 'Big'


# %%
#chargement des données regions 
regions = pd.read_excel("sales.xlsx", "Regions", header=None, skiprows=1,
                        names=['regionID', 'suburb', 'city', 'cp', 'Longitude', 'Latitude', 'address'])
region = regions[['regionID', 'cp', 'city']]



# %%
#chargement des données products 
products = pd.read_excel("sales.xlsx", "Products", header=None, skiprows=1,
                         names=['productID', 'productName'])


# %%
#chargement des données sales
sales = pd.read_excel("sales.xlsx", "Sales Orders", header=None, skiprows=1,
                      names=['orderID', 'orderDate', 'Ship Date', 'CustomerID', 'Channel', 'Currency Code', 'Warehouse Code', 'regionID', 'productID', 
                             'quantity', 'unitprice', 'unitCost'])


# %%
# Convertir 'orderDate' au format datetime
sales['orderDate'] = pd.to_datetime(sales['orderDate'], errors='coerce')
sales['Year'] = sales['orderDate'].dt.year
sales['TotalSales'] = sales['quantity'] * sales['unitprice']
# %%
current_year = 2019  
previous_year = 2018 
# %%
current_year_sales = sales[sales['Year'] == current_year]
previous_year_sales = sales[sales['Year'] == previous_year]
# %%
current_year_product_sales = current_year_sales.groupby('productID')['TotalSales'].sum().reset_index()
previous_year_product_sales = previous_year_sales.groupby('productID')['TotalSales'].sum().reset_index()
# %%
product_sales_comparison = pd.merge(current_year_product_sales, previous_year_product_sales, on='productID', suffixes=('_CY', '_PY'))
product_sales_comparison = product_sales_comparison.merge(products[['productID', 'productName']], on='productID')
# %%
print("Données de comparaison des ventes par produit :")
print(product_sales_comparison)


# %%
#Sales CY vs Sales PY by Product:
plt.figure(figsize=(14, 7))
sns.barplot(data=product_sales_comparison, x='productName', y='TotalSales_CY', label='Ventes CY', color='b', alpha=0.6)
sns.barplot(data=product_sales_comparison, x='productName', y='TotalSales_PY', label='Ventes PY', color='r', alpha=0.6)
plt.xlabel('Nom du Produit')
plt.ylabel('Total des Ventes')
plt.title('Comparaison des ventes CY vs PY par produit')
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()
# %%

# Convertir 'orderDate' en format datetime
sales['orderDate'] = pd.to_datetime(sales['orderDate'], errors='coerce')
sales['Month'] = sales['orderDate'].dt.month
sales['Year'] = sales['orderDate'].dt.year
sales['TotalSales'] = sales['quantity'] * sales['unitprice']
# %%
# Définir les années 
current_year = 2019  
previous_year = 2018  
# %%
current_year_sales = sales[sales['Year'] == current_year]
previous_year_sales = sales[sales['Year'] == previous_year]
# %%
current_year_monthly_sales = current_year_sales.groupby('Month')['TotalSales'].sum().reset_index()
previous_year_monthly_sales = previous_year_sales.groupby('Month')['TotalSales'].sum().reset_index()

# %%
monthly_sales_comparison = pd.merge(current_year_monthly_sales, previous_year_monthly_sales, on='Month', suffixes=('_CY', '_PY'))

# %%
# Vérifier les données 
print("Données de comparaison des ventes par mois :")
print(monthly_sales_comparison)
# %%
#Sales CY vs Sales PY by Month
plt.figure(figsize=(14, 7))
sns.barplot(data=monthly_sales_comparison, x='Month', y='TotalSales_CY', color='b', alpha=0.6, label='Ventes CY')

sns.barplot(data=monthly_sales_comparison, x='Month', y='TotalSales_PY', color='r', alpha=0.6, label='Ventes PY')

plt.xlabel('Mois')
plt.ylabel('Total des Ventes')
plt.title('Comparaison des ventes CY vs PY par mois')

month_labels = pd.date_range(start='2024-01-01', periods=12, freq='M').strftime('%B')

plt.xticks(ticks=np.arange(12), labels=month_labels, rotation=45)
plt.legend()
plt.tight_layout()
plt.show()

# %%

# Sales et Sales PY par Customer Names

current_year_customer_sales = current_year_sales.groupby('CustomerID')['TotalSales'].sum().reset_index()
previous_year_customer_sales = previous_year_sales.groupby('CustomerID')['TotalSales'].sum().reset_index()

# %%
current_year_customer_sales = current_year_customer_sales.merge(customers[['customerID', 'customerName']], left_on='CustomerID', right_on='customerID')
previous_year_customer_sales = previous_year_customer_sales.merge(customers[['customerID', 'customerName']], left_on='CustomerID', right_on='customerID')
# %%
customer_sales_comparison = pd.merge(current_year_customer_sales[['customerName', 'TotalSales']], 
                                     previous_year_customer_sales[['customerName', 'TotalSales']],
                                     on='customerName', suffixes=('_CY', '_PY'))
# %%
# Vérifier les données 
print("Données de comparaison des ventes par client :")
print(customer_sales_comparison)

# %%

plt.figure(figsize=(14, 8))


bar_width = 0.35 
index = np.arange(len(customer_sales_comparison)) 

plt.barh(index - bar_width / 2, customer_sales_comparison['TotalSales_CY'], bar_width, label='Ventes CY', color='green')

plt.barh(index + bar_width / 2, customer_sales_comparison['TotalSales_PY'], bar_width, label='Ventes PY', color='black')

plt.xlabel('Total des Ventes')
plt.ylabel('Nom du Client')
plt.title('Comparaison des ventes CY vs PY par client')
plt.yticks(index, customer_sales_comparison['customerName'])
plt.legend()

plt.axvline(0, color='black', linewidth=2)

plt.tight_layout()
plt.show()
# %%



#Sales by City. 

sales['orderDate'] = pd.to_datetime(sales['orderDate'], errors='coerce')
sales['TotalSales'] = sales['quantity'] * sales['unitprice']


city_sales = sales.groupby('regionID')['TotalSales'].sum().reset_index()
city_sales = city_sales.merge(regions[['regionID', 'city']], on='regionID')

top_cities = city_sales.nlargest(5, 'TotalSales')

plt.figure(figsize=(10, 8))
plt.pie(top_cities['TotalSales'], labels=top_cities['city'], autopct='%1.1f%%', startangle=140, colors=sns.color_palette('pastel'), 
        wedgeprops=dict(width=0.3)) 
plt.title('Top 5 Villes par Ventes')
plt.axis('equal')
plt.show()

# %%



#Profit, Profit PY et Profit Margin par Channel
sales_comparison = pd.merge(current_year_sales, previous_year_sales, on='productID', suffixes=('_CY', '_PY'))
sales_comparison = sales_comparison.merge(products[['productID', 'productName']], on='productID')

print("\nSales Comparison Data:")
print(sales_comparison)

plt.figure(figsize=(14, 7))
sns.barplot(data=sales_comparison, x='productName', y='TotalSales_CY', label='Ventes CY', color='b', alpha=0.6)
sns.barplot(data=sales_comparison, x='productName', y='TotalSales_PY', label='Ventes PY', color='r', alpha=0.6)
plt.xlabel('Nom du Produit')
plt.ylabel('Total des Ventes')
plt.title('Comparaison des ventes CY vs PY par produit')
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()
# %%




#Sales et Sales PY par Customer Names:

data = {
    'CustomerName': ['Customer A', 'Customer B', 'Customer C'],
    'TotalSales_CY': [150000, 200000, 170000],
    'TotalSales_PY': [140000, 180000, 160000]
}
# %%
customer_sales = pd.DataFrame(data)

# %%
plt.figure(figsize=(14, 8))


bar_width = 0.35 
index = np.arange(len(customer_sales)) 

plt.barh(index - bar_width / 2, customer_sales['TotalSales_CY'], bar_width, label='Ventes CY', color='green')

plt.barh(index + bar_width / 2, customer_sales['TotalSales_PY'], bar_width, label='Ventes PY', color='black')

plt.xlabel('Total des Ventes')
plt.ylabel('Nom du Client')
plt.title('Comparaison des ventes CY vs PY par client')
plt.yticks(index, customer_sales['CustomerName'])
plt.legend()

plt.axvline(0, color='black', linewidth=2)

plt.tight_layout()
plt.show()
# %%
