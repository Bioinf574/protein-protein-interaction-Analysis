from rdkit import Chem
print(Chem.MolFromSmiles("COOH"))
print(Chem.MolToSmiles(Chem.MolFromSmiles("COOH")))