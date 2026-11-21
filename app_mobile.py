def auditar_repousos_sequencia(linhas_simuladas, sede="COB"):
    """Valida descanso entre o fim de uma OS e o início da OS seguinte (14h Sede | 9h Fora)."""
    alertas_rep = []
    # Ordena por OS e junta por blocos de dias efetivos para comparar transições
    os_ordenadas = sorted(linhas_simuladas, key=lambda r: (int(r["OS"]), str(r["Dias"])))
    for i in range(len(os_ordenadas) - 1):
        atual = os_ordenadas[i]
        seguinte = os_ordenadas[i + 1]
        if atual["OS"] == seguinte["OS"]:
            continue
        m_ret = hora_para_min(atual["HoraRet"])
        m_apr = hora_para_min(seguinte["HoraApr"])
        if m_ret is None or m_apr is None:
            continue

        descanso = (m_apr - m_ret) if m_apr >= m_ret else ((24 * 60 - m_ret) + m_apr)
        local_desc = str(atual["LocalRet"]).strip().upper()
        em_sede = (local_desc == sede.upper()) or bool(atual.get("RepSede", False))
        min_exigido = 14 * 60 if em_sede else 9 * 60
        tipo_local = "Sede (14h00)" if em_sede else "Fora da Sede (09h00)"

        if descanso < min_exigido:
            alertas_rep.append(
                f"🔴 OS {atual['OS']:03d} → OS {seguinte['OS']:03d} ({local_desc}): "
                f"Repouso insuficiente de {fmt_hm(descanso)} [Mín. {tipo_local}]"
            )
        else:
            alertas_rep.append(
                f"🟢 OS {atual['OS']:03d} → OS {seguinte['OS']:03d} ({local_desc}): "
                f"Repouso válido de {fmt_hm(descanso)} [>= {tipo_local}]"
            )
    return alertas_rep
with aba_edit:
    st.markdown(
        "Edita as **OS pretendidas** usando os números dos comboios (ex: `722 - (TR) - 4674`). "
        "Marca a caixa **Simular** nas linhas que queres auditar (**Repouso Sede 14h**, **Fora 9h**, **TR 45m** e **PNT 6h-9h**)."
    )

    with st.expander("🚆 Gerir Horários de Comboios (comboios.csv)", expanded=False):
        c_g1, c_g2, c_g3, c_g4, c_g5 = st.columns(5)
        n_cb = c_g1.text_input("Comboio", placeholder="722")
        n_or = c_g2.text_input("Origem", placeholder="AVO")
        n_pa = c_g3.text_input("Partida", placeholder="15:56")
        n_de = c_g4.text_input("Destino", placeholder="COB")
        n_ch = c_g5.text_input("Chegada", placeholder="16:30")
        if st.button("➕ Guardar Comboio na Sessão"):
            if n_cb.strip() and n_pa.strip() and n_ch.strip():
                st.session_state["base_comboios"][n_cb.strip()] = {
                    "Origem": n_or.strip() or "EST",
                    "Partida": n_pa.strip(),
                    "Destino": n_de.strip() or "EST",
                    "Chegada": n_ch.strip()
                }
                st.success(f"Comboio {n_cb} adicionado!")

        pdf_hor = st.file_uploader("Ou extrair comboios de um PDF de Horários CP:", type=["pdf"], key="pdf_hor")
        if pdf_hor and st.button("📥 Extrair Horários do PDF"):
            ext = extrair_horarios_de_pdf_cp(pdf_hor)
            st.session_state["base_comboios"].update(ext)
            st.success(f"Extraídos {len(ext)} comboios do PDF!")

    arq_sim = st.file_uploader("Carrega a Escala Base para Editar (.xlsx ou .pdf)", type=["xlsx", "pdf"], key="sim")
    if arq_sim:
        d_sim, meta_sim = ler_ficheiro_carregado(arq_sim)
        arq_sim.seek(0)
        tot_os_sim = obter_total_os(d_sim)
        sede_sim = meta_sim.get("Sede", "COB")

        os_sel = st.multiselect(
            "Seleciona as OS que queres visualizar/editar:",
            options=list(range(1, tot_os_sim + 1)),
            default=[min(117, tot_os_sim)]
        )

        if os_sel:
            linhas_orig = [x for x in d_sim if x["OS"] in os_sel]
            df_edit_base = pd.DataFrame(linhas_orig)[
                ["OS", "Dias", "LocalApr", "HoraApr", "Serv", "LocalRet", "HoraRet", "PNT", "RepFora"]
            ].copy()

            # Coluna de visto "Simular" no início da tabela
            df_edit_base.insert(0, "Simular", True)

            st.caption("✏️ Edita a coluna **Serv** (ou horários) e deixa o visto **Simular** ativo apenas nas linhas a auditar:")
            df_editado = st.data_editor(
                df_edit_base,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "Simular": st.column_config.CheckboxColumn("Simular", help="Marcar para incluir esta linha na auditoria", default=True),
                    "OS": st.column_config.NumberColumn("OS", min_value=1, max_value=365, step=1, format="%d"),
                    "Dias": st.column_config.TextColumn("Dias"),
                    "LocalApr": st.column_config.TextColumn("Local Apr"),
                    "HoraApr": st.column_config.TextColumn("Hora Apr"),
                    "Serv": st.column_config.TextColumn("Designação dos Serviços", width="large"),
                    "LocalRet": st.column_config.TextColumn("Local Ret"),
                    "HoraRet": st.column_config.TextColumn("Hora Ret"),
                    "PNT": st.column_config.TextColumn("PNT"),
                    "RepFora": st.column_config.CheckboxColumn("Rep Fora")
                },
                key="editor_os_sim"
            )

            col_b1, col_b2 = st.columns([2, 1])
            recalc_auto = col_b2.checkbox("Recalcular horários via comboios.csv", value=True)

            if col_b1.button("🧪 Simular e Auditar Linhas Selecionadas", type="primary", use_container_width=True):
                df_selecionadas = df_editado[df_editado["Simular"] == True].copy()

                if df_selecionadas.empty:
                    st.warning("⚠️ Marca pelo menos uma linha com o visto **Simular** para correr a auditoria.")
                else:
                    infracoes_totais = []
                    linhas_comp_pdf = []
                    linhas_para_repouso = []

                    for idx, row in df_selecionadas.iterrows():
                        os_num = int(row["OS"])
                        serv_txt = str(row["Serv"]).strip()
                        dias_txt = str(row["Dias"]).strip()

                        loc_a, hora_a, loc_r, hora_r, pnt_calc, alertas_linha = auditar_linha_editada(
                            serv_txt, st.session_state["base_comboios"], sede=sede_sim
                        )

                        # Se o utilizador optar por manter horários manuais quando não há comboio no CSV
                        if not recalc_auto or not hora_a:
                            loc_a = str(row["LocalApr"]).strip() or loc_a
                            hora_a = str(row["HoraApr"]).strip() or hora_a
                            loc_r = str(row["LocalRet"]).strip() or loc_r
                            hora_r = str(row["HoraRet"]).strip() or hora_r
                            pnt_calc = str(row["PNT"]).strip() or pnt_calc

                        # Recalcula PNT se hora_a e hora_r estiverem preenchidas manualmente
                        ma, mr = hora_para_min(hora_a), hora_para_min(hora_r)
                        if ma is not None and mr is not None and not recalc_auto:
                            dur_m = (mr - ma) if mr >= ma else ((24 * 60 - ma) + mr)
                            pnt_calc = f"{dur_m // 60:02d}:{dur_m % 60:02d}"
                            if dur_m < 6 * 60:
                                alertas_linha.append(f"🔴 PNT abaixo do mínimo (06h00): {fmt_hm(dur_m)}")
                            elif dur_m > 9 * 60:
                                alertas_linha.append(f"🔴 PNT acima do máximo (09h00): {fmt_hm(dur_m)}")

                        for al in alertas_linha:
                            infracoes_totais.append(f"OS {os_num:03d} [{dias_txt}]: {al}")

                        linhas_para_repouso.append({
                            "OS": os_num,
                            "Dias": dias_txt,
                            "LocalApr": loc_a,
                            "HoraApr": hora_a,
                            "LocalRet": loc_r,
                            "HoraRet": hora_r,
                            "RepFora": bool(row["RepFora"])
                        })

                        # Texto original vs simulado para o PDF
                        orig_match = next((o for o in linhas_orig if o["OS"] == os_num and o["Dias"] == dias_txt), None)
                        txt_orig = f"[{orig_match['Dias']}] {orig_match['HoraApr']}-{orig_match['HoraRet']} | {orig_match['Serv']} ({orig_match['PNT']})" if orig_match else "Nova linha simulada"
                        txt_edit = f"[{dias_txt}] {hora_a}-{hora_r} | {serv_txt} ({pnt_calc})"
                        linhas_comp_pdf.append((f"OS {os_num:03d}", txt_orig, txt_edit))

                    # Audita repousos consecutivos entre as OS marcadas (14h Sede / 9h Fora)
                    alertas_descanso = auditar_repousos_sequencia(linhas_para_repouso, sede=sede_sim)
                    infracoes_totais.extend(alertas_descanso)

                    st.subheader("📋 Resultado da Auditoria (Linhas Marcadas)")
                    if not infracoes_totais:
                        st.success("✅ Nenhuma infração detetada nas linhas selecionadas!")
                    else:
                        for item in infracoes_totais:
                            if item.startswith("🔴"):
                                st.error(item)
                            elif item.startswith("⚠️"):
                                st.warning(item)
                            else:
                                st.success(item)

                    # Tabela resumo recalculada das linhas simuladas
                    df_resumo_sim = pd.DataFrame(linhas_para_repouso)
                    st.dataframe(df_resumo_sim, use_container_width=True, hide_index=True)

                    pdf_sim_bytes = gerar_pdf_impacto_simulacao(
                        linhas_comp_pdf,
                        infracoes_totais or ["Sem infrações detetadas."],
                        titulo_esc=meta_sim.get("Escala", "COBR20")
                    )
                    st.download_button(
                        "📄 Descarregar PDF de Impacto da Simulação",
                        data=pdf_sim_bytes,
                        file_name=f"Simulacao_{meta_sim.get('Escala', 'ORV')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
