# ── 复合规则：商家(old) + 商品 双条件，优先级高于简单规则 ──
if not compound_rules.empty:
    bill_product = df['商品'].astype(str)

    for _, r in compound_rules.iterrows():
        cp = str(r['商家(old)']).strip()
        pp = str(r['商品']).strip()
        if not cp or cp.lower() == 'nan' or not pp or pp.lower() == 'nan':
            continue

        use_regex = int(r.get('is_regex', 0)) == 1
        try:
            hit = (
                df['商家(old)'].str.contains(cp, regex=use_regex, na=False) &
                bill_product.str.contains(pp, regex=use_regex, na=False)
            )
            if not hit.any():
                continue

            for c in rename_map:
                val = r.get(c)
                if not (pd.notna(val) and str(val).strip() not in ('', 'nan')):
                    continue

                val_str = str(val).strip()

                if c == '描述':
                    # ✅ 修复①：描述走 描述_rule，交给 construct_description 统一处理
                    df.loc[hit, '描述_rule'] = val_str

                elif c == '商家':
                    # ✅ 升级②：商家推导优先级控制
                    # 简单规则已命中（商家非空且非交易对方原值）→ 不覆盖
                    # 简单规则未命中 → 用复合规则的商家
                    simple_hit_mask = hit & (
                        df['商家'].notna() &
                        (df['商家'].astype(str).str.strip() != '') &
                        (df['商家'].astype(str).str.strip() !=
                         df['商家(old)'].astype(str).str.strip())
                    )
                    compound_only_mask = hit & ~simple_hit_mask
                    if compound_only_mask.any():
                        df.loc[compound_only_mask, '商家'] = val_str

                else:
                    # 其他列（子类别、主类别等）直接写入，保持原有行为
                    df.loc[hit, c] = val_str

            logger.info(f"复合规则 [{cp} & 商品含'{pp}']: {hit.sum()} 条")

        except re.error as e:
            logger.warning(f"复合规则正则错误 '{cp}': {e}")
