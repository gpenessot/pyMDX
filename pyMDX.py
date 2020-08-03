# Librairies

import clr # > pip install pythonnet
from pathlib import Path

root = Path(r'C:\Windows\Microsoft.NET\assembly\GAC_MSIL')
adomd_path = str(max((root / 'Microsoft.AnalysisServices.AdomdClient').iterdir())
                           / 'Microsoft.AnalysisServices.AdomdClient.dll')

clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference(adomd_path)

import System
from System.Data import DataTable
import Microsoft.AnalysisServices.AdomdClient as ADOMD

def ExtractionDonnees(mdx_query, conn_string):
    """
    Cette fonction permet d'extraire les données d'un cube.
    Les arguments sont :
    - la chaîne de connexion au cube
    - la requête MDX
    La fonction retourne une table
    """
    
    dataadapter = ADOMD.AdomdDataAdapter(mdx_query, conn_string)
    table = DataTable()
    dataadapter.Fill(table)
    
    return table

def NettoyageNomsCol(df) :
    """
    Cette fonction nettoie le nom des dimensions issues du cube.
    L'argument est le nom du dataframe.
    
    Exemple :
    > df.columns
    > Index(['[Informations - Compte].[SIREN - Compte].[SIREN - Compte].[MEMBER_CAPTION]',
      dtype='object')
    > NettoyageNomsCol(df)
    > df.columns
    > Index('SIREN - Compte',
    dtype='object')
    """
    df.columns = [x.split('.')[1].strip('[]') for x in df.columns]
    
    return df
