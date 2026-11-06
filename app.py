import streamlit as st
import json
import re
import io
from google import genai
from google.genai import types
from pptx import Presentation

st.set_page_config(page_title="AIスライド生成ツール", page_icon="📊")

st.title("📊 AIスライド自動生成ツール")
st.write("プロット（原案）を入力するだけで、ルールに沿ったPowerPointファイルを作成します。")

# APIキーの読み込み（Secretsから取得）
try:
    api_key = st.secrets["GEMINI_API_KEY"]
except KeyError:
    st.error("APIキーが設定されていません。StreamlitのSecretsを設定してください。")
    st.stop()

# ルールの設定（裏側で固定）
SLIDE_RULES = """
・1枚1テーマ。タイトルだけを読んでも話がつながる構成にする。
・「現状と課題：」のような分類名を前置きせず、タイトルに結論（主張）を書く。
・「3つの理由」のような項目数で済ませず、中身の構造や結論を書く。
・箇条書きは体言止めを避け、意味が通じる文章にする。
"""

# プロットの入力
slide_plot = st.text_area("スライドにしたい原案（プロット）を入力してください", height=200)

# 実行ボタン
if st.button("スライドを生成する"):
    if not slide_plot:
        st.error("原案を入力してください。")
    else:
        with st.spinner("AIがスライドを生成中です...（10秒ほどお待ちください）"):
            try:
                # Geminiの準備
                client = genai.Client(api_key=api_key)
                system_instruction = f"あなたはプロの戦略コンサルタントです。以下の【スライド作成規約】を厳格に守って資料を構成してください。\n\n【スライド作成規約】\n{SLIDE_RULES}"
                prompt = f"""
                以下の【プロット（原案）】をもとに、論理的なスライド構成を作成してください。
                【プロット（原案）】\n{slide_plot}
                出力は、必ず以下のJSON形式の配列のみとしてください。Markdownの装飾は不要です。
                [
                  {{"title": "1枚目のタイトル", "content": ["ポイント1", "ポイント2"]}}
                ]
                """

                # AIによる生成（安定の2.0 Flashモデルを使用）
                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=system_instruction,
                        temperature=0.2
                    )
                )

                # データの抽出
                json_str = re.search(r'\[.*\]', response.text, re.DOTALL)
                slide_data = json.loads(json_str.group(0)) if json_str else json.loads(response.text)

                # PPTXの作成
                prs = Presentation()
                for slide_info in slide_data:
                    slide_layout = prs.slide_layouts[1] 
                    slide = prs.slides.add_slide(slide_layout)
                    slide.shapes.title.text = slide_info["title"]
                    tf = slide.shapes.placeholders[1].text_frame
                    if len(slide_info["content"]) > 0:
                        tf.text = slide_info["content"][0]
                        for point in slide_info["content"][1:]:
                            p = tf.add_paragraph()
                            p.text = point

                # メモリ上にファイルを保存（PCにゴミを残さない）
                pptx_io = io.BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)

                st.success("スライドの生成が完了しました！下のボタンからダウンロードしてください。")
                
                # ダウンロードボタンの表示
                st.download_button(
                    label="📥 PowerPointファイルをダウンロード",
                    data=pptx_io,
                    file_name="Generated_Slides.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
