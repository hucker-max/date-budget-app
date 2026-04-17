on find_placeholder_shape(the_slide, placeholder_kind)
	tell application "Microsoft PowerPoint"
		repeat with sh in shapes of the_slide
			try
				if placeholder type of sh is placeholder_kind then return sh
			end try
		end repeat
	end tell
	return missing value
end find_placeholder_shape

on set_slide_text(the_slide, title_text, body_text)
	tell application "Microsoft PowerPoint"
		set title_shape to my find_placeholder_shape(the_slide, placeholder type title placeholder)
		if title_shape is missing value then set title_shape to my find_placeholder_shape(the_slide, placeholder type center title placeholder)
		if title_shape is not missing value then set content of text range of text frame of title_shape to title_text
		
		set body_shape to my find_placeholder_shape(the_slide, placeholder type body placeholder)
		if body_shape is missing value then set body_shape to my find_placeholder_shape(the_slide, placeholder type subtitle placeholder)
		if body_shape is not missing value then set content of text range of text frame of body_shape to body_text
	end tell
end set_slide_text

on add_text_slide(the_pres, slide_title, slide_body)
	tell application "Microsoft PowerPoint"
		set new_slide to make new slide at the end of slides of the_pres with properties {layout:slide layout text slide}
		my set_slide_text(new_slide, slide_title, slide_body)
	end tell
end add_text_slide

set output_dir to POSIX path of "/Users/okadaigo/Documents/New project/"
set pptx_path to output_dir & "PEC1-4_材料課題_簡易まとめ.pptx"

set slide2_body to "• PEC-1: ネットワーク側の光化" & return & "  データセンター間・サーバ間を低遅延・大容量・低消費電力でつなぎたい" & return & return & "• PEC-2: ボード間の光化" & return & "  Switch ASIC 近傍まで optical engine を寄せて、装置内の通信電力を下げたい" & return & return & "• PEC-3: パッケージ間の光化" & return & "  半導体パッケージ間を高密度・低消費電力でつなぎたい" & return & return & "• PEC-4: ダイ間・チップ内の光化" & return & "  die-to-die / intra-chip まで光を入れて、計算機内部構造を変えたい"

set slide3_body to "• PEC-1 でやりたいこと" & return & "  長距離伝送を低損失・大容量で成立させる" & return & "• 想定材料" & return & "  石英ファイバ、シリカ PLC / AWG、Si フォトニクス、InP 系" & return & "• 材料課題" & return & "  伝送損失、波長分離、耐湿・耐熱、長期安定性" & return & return & "• PEC-2 でやりたいこと" & return & "  ボード間配線を光化し、装置内通信の電力を下げる" & return & "• 想定材料" & return & "  Si / SiO2 / InP / InGaAlAs、SiOx 導波路、コネクタ・接着材" & return & "• 材料課題" & return & "  熱、接合、アライメント、コネクタ小型化、近接実装"

set slide4_body to "• PEC-3 でやりたいこと" & return & "  package-to-package 光接続を成立させる" & return & return & "• 材料スタックの見方" & return & "  Si フォトニクス光回路 / InP 系膜型活性層 / クラッド / 放熱層 / 接着・封止材" & return & return & "• 材料課題" & return & "  熱で光結合が変わる" & return & "  反りでアライメントがずれる" & return & "  1310/1550nm 損失が効く" & return & "  リフロー・熱サイクル・量産性が支配的" & return & return & "• 研究の主戦場" & return & "  熱・反り・光損失・信頼性を同時に見る必要がある"

set slide5_body to "• PEC-4 でやりたいこと" & return & "  die-to-die / chip 内 optical layer を成立させる" & return & return & "• 仮説的な材料スタック" & return & "  Si / SiN 導波路、InP 薄膜 or Ge / Si、低k / クラッド、Cu / SiCN bonding、熱拡散材" & return & return & "• 材料課題" & return & "  bonding の位置ずれ・接合信頼性" & return & "  optical layer と electrical layer の熱干渉" & return & "  長期信頼性、応力、封止" & return & return & "• 見立て" & return & "  全面光化より、I/O タイル起点の部分光化が自然そう"

set slide6_body to "• OPE" & return & "  高温耐性・CTE は比較的強い" & return & "  ただし 1310/1550nm の導波路損失が不明" & return & return & "• BT 樹脂系" & return & "  基板材料として有望" & return & "  ただし反り・剥離・光学損失データが不足" & return & return & "• ユピゼータEP" & return & "  光学部材・光学樹脂候補" & return & "  ただし耐熱・近赤外・CTE の定量データが不足"

set slide7_body to "• 技術としてやりたいこと" & return & "  光電融合を、よりコンピュータの近くまで持ち込む" & return & return & "• 材料研究としての問い" & return & "  どの材料が PEC-3 のどの役割を担えるか" & return & "  何のデータが足りないと最終判断できないか" & return & return & "• 今後の評価軸" & return & "  260℃リフロー耐性 / 1310・1550nm 損失 / CTE・寸法安定性" & return & return & "• 次にやること" & return & "  公開エビデンス比較と、評価項目-測定方法の対応づけ"

tell application "Microsoft PowerPoint"
	activate
	set pres to make new presentation
	set s1 to slide 1 of pres
	set layout of s1 to slide layout title slide
	my set_slide_text(s1, "PEC1-4のリサーチ結果と材料課題", "卒業研究向け 簡易まとめ" & return & "2026-04-16")
	
	my add_text_slide(pres, "1. PEC1-4で技術として何をしたいか", slide2_body)
	my add_text_slide(pres, "2. PEC-1 / PEC-2", slide3_body)
	my add_text_slide(pres, "3. PEC-3", slide4_body)
	my add_text_slide(pres, "4. PEC-4", slide5_body)
	my add_text_slide(pres, "5. MGC材料への引きつけ", slide6_body)
	my add_text_slide(pres, "6. 研究としての着地点", slide7_body)
	
	save pres in POSIX file pptx_path as save as Open XML presentation
end tell
