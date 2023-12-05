import pandas as pd

main = pd.read_excel("pm.xlsx", sheet_name="main")
state = pd.read_excel("pm.xlsx", sheet_name="state")

output_name = str(input("Enter file name i.e. Name.xlsx: "))

vlookup = pd.merge(main, state, on="Name", how="inner")

pivottable = vlookup.pivot_table(index=["Name", "Age"], values="State", aggfunc="count").reset_index()
pivottable.columns = ["Name", "Age", "Count of State"]

pivottable.to_excel(output_name, sheet_name="Pivot Table", index=False)

pivottable2 =vlookup.pivot_table(index="State", values="Name", aggfunc="count").reset_index()
pivottable2.columns = ["State", "Count of Name"]

with pd.ExcelWriter(output_name, engine="openpyxl", mode="a") as writer:
   pivottable2.to_excel(writer, sheet_name="State Vs Count of Names", index=False)

print("File is Created")
