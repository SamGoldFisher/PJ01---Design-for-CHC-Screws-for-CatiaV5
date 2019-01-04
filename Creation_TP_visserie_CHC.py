#coding:utf-8
# Préparation de l'environnement pour excel
import xlrd

# Ouverture du classeur
l= xlrd.open_workbook("L1.xls")
k = xlrd.open_workbook("K.xls")
biblio = open("biblio.xls", "w")

# Récupération des premières feuilles
L0 = l.sheet_by_name(l.sheet_names()[0])
K0 = k.sheet_by_name(k.sheet_names()[0])

# Calcul du nombre de longueur de vis
nb_longueur=L0.ncols;
print("le nombre de longueur de vis est : ",nb_longueur)

# Calcul du nombre de colonne que l'on doit créer (en incluant 'PartNumber' et L(mm))
nb_para=K0.ncols;
nb_para+=2;
print("le nombre de paramètres dans la table de paramétrage est :",nb_para)

# Calcul du nombre de diametre différent a traiter
nb_dia=K0.nrows-1;
print("le nombre de diamètre de vis à traiter est :",nb_dia)

# Calcul du nombre de référence à créer soit le nombre de lignes
nb_ref=nb_dia*nb_longueur;
print("le nombre de références de vis à créer est :",nb_ref)

# Création du tableau
indice=0;
for i in range (0,nb_dia): #concerne le diametre 
	d=diametre=K0.cell_value(i+1,0)
	if d!=2.5:
		d=round(d)
		diametre=d
	else:
		d=d
	g=K0.cell_value(i+1,1)
	k1=K0.cell_value(i+1,2)
	p=K0.cell_value(i+1,3)
	print ("le diametre traité est : ",diametre)
	for j in range (0, nb_longueur-1): #concerne la longueur de vis créée pour un diamètre
		if indice==0:
			biblio.write("PartNumber \tD (mm)\t L(mm)\tg (mm) \tk1 (mm)\tp (mm)\n")
			indice=1;
		else:
			longueur=round((L0.cell_value(0,j)))
			# Ecriture de la du part number:
			biblio.write("Vis CHC M")
			biblio.write(str(diametre))
			biblio.write("x")
			biblio.write(str(longueur))
			biblio.write("\t")
			# Ecriture des autres paramètres :
			biblio.write(str(diametre))#D
			biblio.write("\t")
			biblio.write(str(longueur))#longueur
			biblio.write("\t")
			biblio.write(str(g))#g
			biblio.write("\t")
			biblio.write(str(k1))#k1
			biblio.write("\t")
			biblio.write(str(p))#p
			biblio.write("\t")
			biblio.write("\n")
			indice+=1;
	print("\n")
	
 # I=data_L.sheet_names()[0].nrows-1;
 # J=data_L.sheet_names()[0].ncols-1;
# print(I)
# print(J)
# print("ensuite voici le texte écrit dans le fichier donner par ligne puis par colonne :\n" )
# for i in range(0,I+1):
	# if i==0:
			# biblio.write("PartNumber \tD(mm) \tg(mm) \tk1(mm) \tp(mm)\n");		
	# for j in range(0,J+1):
			# a=format(data_L.sheet_names()[0].cell_value(i,j));
			# print(a)
			# biblio.write(str(a));
			# biblio.write("\t");
	# biblio.write("\n");

# Fermeture de la bibliothèque a construire
biblio.close()

import os
os.system("pause")
trash=os.system("cls")