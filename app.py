m5.metric("Não conferidos", summary["nao_conferidos"])
m6.metric("Extras Benner", summary["extras_benner"])

st.divider()

if summary["pendencias"] == 0:
    st.success(
        "Nenhuma falta foi encontrada nos campos que puderam ser conferidos."
    )
else:
    st.error(
        f"Foram encontradas {summary['pendencias']} faltas para conferência."
    )

if summary["nao_conferidos"] > 0:
    st.warning(
        f"{summary['nao_conferidos']} conferências ficaram como 'NÃO CONFERIDO' "
        "porque o relatório do Benner não trouxe todos os campos necessários."
    )

# -------------------------------------------------------------------------
# Filtros
# -------------------------------------------------------------------------
if not detailed_df.empty:
    cargos = st.multiselect(
        "Filtrar cargo",
        sorted(detailed_df["Cargo"].unique()),
    )

    status_filter = st.multiselect(
        "Filtrar status",
        [
            "OK",
            "FALTA",
            "NÃO CONFERIDO — COLUNA AUSENTE",
            "NÃO CONFERIDO — EXAME NÃO MAPEADO",
            "EXTRA NO BENNER",
        ],
        default=["FALTA", "NÃO CONFERIDO — COLUNA AUSENTE"],
    )

    view = detailed_df.copy()

    if cargos:
        view = view[view["Cargo"].isin(cargos)]

    if status_filter:
        view = view[view["Status"].isin(status_filter)]

    st.subheader("🔎 Detalhado — PCMSO × Benner")
    st.dataframe(
        view,
        use_container_width=True,
        hide_index=True,
    )

st.subheader("📋 Status por cargo")
st.dataframe(
    role_status_df,
    use_container_width=True,
    hide_index=True,
)

with st.expander("⚠️ Pendências / não conferidos"):
    if pending_df.empty:
        st.success("Nenhuma pendência.")
    else:
        st.dataframe(
            pending_df,
            use_container_width=True,
            hide_index=True,
        )

with st.expander("➕ Extras cadastrados no Benner"):
    if extras_df.empty:
        st.write("Nenhum extra foi identificado.")
    else:
        st.dataframe(
            extras_df,
            use_container_width=True,
            hide_index=True,
        )

with st.expander("❌ Cargos do PCMSO não encontrados no Benner"):
    if missing_roles_df.empty:
        st.success("Todos os cargos do PCMSO tiveram correspondência.")
    else:
        st.dataframe(
            missing_roles_df,
            use_container_width=True,
            hide_index=True,
        )

excel_bytes = build_excel(
    summary,
    role_status_df,
    detailed_df,
    pending_df,
    extras_df,
    missing_roles_df,
)

st.download_button(
    "⬇️ Baixar auditoria em Excel",
    data=excel_bytes,
    file_name="auditoria_pcmso_x_benner.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)

st.caption(
    "Regra especial: quando o PCMSO marca "
    "'Mud. Risco Ocupacional | Retorno', o aplicativo confere "
    "principalmente o campo 'Mudança de Função'. O campo 'Retorno ao Trabalho' "
    "fica apenas como informação no relatório e não gera falta nesta versão. "
    "Exames compostos, como TGO + TGP, são conferidos por seus componentes."
)
