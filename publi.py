from os import listdir
from os.path import isfile, isdir
import xlrd
# Para poder funcionar la libreria de excel este deve estar en a vercion pip install xlrd==1.2.0 pero no se recomienda ya que este es insegura para servidores.

productos = "productos.xlsx"
openProductos = xlrd.open_workbook(productos)
Productos = openProductos.sheet_by_name("TEMPLETE PARA PRODUCTOS")
datosTienda = openProductos.sheet_by_name("DatosTienda")
helder =""
vendedor = datosTienda.cell_value(0,0)
url = datosTienda.cell_value(0,1)
contadorImagenes = 0
for ip in range(Productos.nrows):
    if ip == 0:
        with open('productos.csv','w') as file_escritura:
            file_escritura.write('Handle,Title,Body (HTML),Vendor,Type,Tags,Published,Option1 Name,Option1 Value,Option2 Name,Option2 Value,Option3 Name,Option3 Value,Variant SKU,Variant Grams,Variant Inventory Tracker,Variant Inventory Qty,Variant Inventory Policy,Variant Fulfillment Service,Variant Price,Variant Compare At Price,Variant Requires Shipping,Variant Taxable,Variant Barcode,Image Src,Image Position,Image Alt Text,Gift Card,SEO Title,SEO Description,Google Shopping / Google Product Category,Google Shopping / Gender,Google Shopping / Age Group,Google Shopping / MPN,Google Shopping / AdWords Grouping,Google Shopping / AdWords Labels,Google Shopping / Condition,Google Shopping / Custom Product,Google Shopping / Custom Label 0,Google Shopping / Custom Label 1,Google Shopping / Custom Label 2,Google Shopping / Custom Label 3,Google Shopping / Custom Label 4,Variant Image,Variant Weight Unit,Variant Tax Code,Cost per item,Status \r')
    if ip != 0:
        vendedorCampo = ""
        tiposCAmpo = ""
        taCampo = ""
        trueCapo = ""
        imgenCampo = ""
        contadorImagenesCampo = ""
        
        if Productos.cell_value(ip,0) != "":
            helder = Productos.cell_value(ip,0).replace(" ", "_")
            vendedorCampo = vendedor
            tiposCAmpo = Productos.cell_value(ip,2)
            taCampo = Productos.cell_value(ip,3)
            trueCapo = "true"
            contadorImagenes = 0
        contadorImagenes = contadorImagenes+1
        if Productos.cell_value(ip,14) != "":
            imgenCampo = f'{url}{Productos.cell_value(ip,14)}'
            contadorImagenesCampo = contadorImagenes

        with open('productos.csv','a', encoding="utf-8") as file_escritura:
            if  Productos.cell_value(ip,5) != "":
                file_escritura.write(f'{helder},{Productos.cell_value(ip,0)},"{Productos.cell_value(ip,1).capitalize()}",{vendedorCampo},"{tiposCAmpo}","{taCampo}",{trueCapo},{Productos.cell_value(ip,4)},{Productos.cell_value(ip,5)},{Productos.cell_value(ip,6)},{Productos.cell_value(ip,7)},{Productos.cell_value(ip,8)},{Productos.cell_value(ip,9)},{str(Productos.cell_value(ip,10))},{Productos.cell_value(ip,11)},shopify,{Productos.cell_value(ip,12)},deny,manual,{Productos.cell_value(ip,13)},,true,true,,{imgenCampo},{contadorImagenesCampo},,false,,"",,,,,,,,,,,,,,{imgenCampo},kg,,,active\r')
            else:
                file_escritura.write(f'{helder},,,,,,,,,,,,,,,,,,,,,,,,{imgenCampo},{contadorImagenesCampo},,,,,,,,,,,,,,,,,,,,,,\r')