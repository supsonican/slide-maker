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

# APIキーの入力（パスワードのように隠れる設定）
api_key = st.text_input("Gemini APIキーを入力してください", type="password")

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
    if not api_key:
        st.error("APIキーを入力してください。")
    elif not slide_plot:
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
                    model='gemini-2.0-flash',
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
</Step>

<Step subtitle="世界に一つだけのURLが発行されます" title="STEP3：Webアプリとして公開する">
1. [Streamlit Community Cloud](https://share.streamlit.io/) にアクセスします。
2. 「**Continue with GitHub**」をクリックし、先ほど作ったGitHubアカウントでログイン・連携を許可します。
3. ログイン後の画面右上にある青いボタン「**Create app**」をクリックします。
4. 「Yep, I have an app」をクリックします。
5. 「Repository」の欄をクリックすると、先ほど作った `slide-maker` が出てくるので選択します。
6. 「Main file path」の欄に `app.py` と入っていることを確認します。
7. 右下の「**Deploy!**」ボタンをクリックします。
</Step>
</Steps>

画面に風船が飛び交うアニメーションが表示されれば、**あなた専用のWebアプリの完成です！**
ブラウザ上部のURL（`https://〜.streamlit.app`）をブックマークしておけば、いつでもスマホや会社のPCからアクセスして、テキストを入れるだけでPPTXを作れるようになります。

<ElicitationsGroup message="設定中に分からないことがあればお知らせください。">
<Elicitation label="GitHubの操作で迷った" query="STEP1のGitHubで、ファイルの作り方や保存（Commit）のやり方がよくわかりませんでした。"/>
<Elicitation label="Streamlitでエラーが出た" query="STEP3でDeployボタンを押したあと、赤いエラー文字が出てアプリが表示されませんでした。"/>
<Elicitation label="アプリの見た目を変えたい" query="無事にWebアプリが完成しました！画面の色を変えたり、ルールのテキストを画面上で編集できるように改造したいです。"/>
</ElicitationsGroup>
