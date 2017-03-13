# Inventory
#### Technologies: ASP, VBScript, HTML, CSS, MS SQL Server
### [All Saints' Catholic Academy] (http://www.allsaints.notts.sch.uk) - Built on 26/11/2013

## Index
* [Installation] (#Install)
* [Usage] (#Usage)
* [Screen Shots] (#Shots)

## Challenege
A web-based Inventory system where users can search for items by their asset tags, description, make, model, supplier, order number, location, purchase date, status, serial number or by clicking on the room image. The system keeps track of all ICT assets with their depreciation values attached. The system can also be used as a stock check facility where the user can scan the asset barcode. Loaned item information can also be attached to an asset.

## <a name="Install">Installation</a>
* To clone the repo
```shell
$ git clone https://github.com/adrianeyre/inventory
$ cd inventory
```

* Set up a web framework such as MS IIS

* Add an ODBC connection to your SQL Server

* Update the file `Connections/PCRoomConnection.asp' with your connection, username and password
```shell
MM_PCRoomConnection_STRING = "dsn=<ODBC Connection>;uid=<USERNAME>;pwd=<PASSWORD>;"
```

## <a name="Shots">Screen Shots</a>
### Search Screen
[![Screenshot](https://raw.githubusercontent.com/adrianeyre/inventory/master/images/screenshot1.png)](https://raw.githubusercontent.com/adrianeyre/inventory/master/images/screenshot1.png "Screen Shot 1")

### Room Run-Down Screen
[![Screenshot](https://raw.githubusercontent.com/adrianeyre/inventory/master/images/screenshot2.png)](https://raw.githubusercontent.com/adrianeyre/inventory/master/images/screenshot2.png "Screen Shot 2")

### Asset Screen
[![Screenshot](https://raw.githubusercontent.com/adrianeyre/inventory/master/images/screenshot3.png)](https://raw.githubusercontent.com/adrianeyre/inventory/master/images/screenshot3.png "Screen Shot 3")
