import pandas as pd

if __name__ == '__main__':
    # Lê o excel e transforma em dataframe
    result_df = pd.read_excel("Results.xlsx")
    site_list_df = pd.read_excel("SiteList.xlsx")

    # Merge entre os dataframes pelo id
    merge_df = pd.merge(site_list_df, result_df, on=["Site ID", "Site Name", "Year", "Equipment"])

    # Filtra apenas os sites de 2023
    merge_df = merge_df.loc[merge_df["Year"] == 2023]

    # Renomeia as colunas
    merge_df = merge_df.rename(columns={
        "Site Name": "Site",
        "State": "Estado",
        "Equipment": "Equipamento",
    })

    # Pega apenas as colunas requisitadas
    merge_df = merge_df[["Site", "Site ID", "Estado", "Equipamento", "Signal (%)", "Quality (0-10)", "Mbps"]]

    # Organiza em ordem alfabética em relação ao estado
    merge_df = merge_df.sort_values("Estado")

    # Exporta o resultado após o tratamento
    merge_df.to_excel("RelatorioFinal.xlsx", index=False)

    # Site com alerta ativos
    alerta_qty = result_df[result_df["Alerts"] == "Yes"].shape[0]
    print(f"Site com alerta ativos: {alerta_qty}")

    # Sites com 0 de qualidade
    qualidade_qty = result_df[result_df["Quality (0-10)"] == 0].shape[0]
    print(f"Sites com 0 de qualidade: {qualidade_qty}")

    # Sites com mais de 80mbps
    site_80mbps_qty = result_df[result_df["Mbps"] > 80].shape[0]
    print(f"Sites com mais de 80mbps: {site_80mbps_qty}")

    # Sites com menos de 10mbps
    site_10mbps_qty = result_df[result_df["Mbps"] < 10].shape[0]
    print(f"Sites com menos de 10mbps: {site_10mbps_qty}")

    # Sites que não estão presentes no result
    site_not_result = pd.merge(
        site_list_df,
        result_df,
        on=["Site ID", "Site Name", "Year", "Equipment"],
        indicator=True
    )
    site_not_result = site_not_result[site_not_result["_merge"] != "both"].shape[0]
    print(f"Sites que não estão presentes no result: {site_not_result}")
