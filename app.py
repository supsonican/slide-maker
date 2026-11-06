import streamlit as st
import json
import re
import io
from google import genai
from google.genai import types
from pptx import Presentation

st.set_page_config(page_title="AIスライド生成ツール", page_icon="📊")
st.title("📊 AIスライド自動生成ツール (Jinbaリレー方式)")

# APIキーの読み込み
try:
    api_key = st.secrets["GEMINI_API_KEY"]
except KeyError:
    st.error("APIキーが設定されていません。StreamlitのSettings > Secretsを確認してください。")
    st.stop()

# ユーザーからの入力
slide_plot = st.text_area("スライドにしたい原案（プロット）を入力してください", height=200)

if st.button("スライドを生成する"):
    if not slide_plot:
        st.error("原案を入力してください。")
    else:
        client = genai.Client(api_key=api_key)
        
        try:
            # ==========================================
            # ステップ1：ストーリーと構成の作成（論理担当）
            # ==========================================
            with st.spinner("ステップ1/3：AIが原案からスライドの論理構成を組み立てています..."):
                step1_instruction = """
                あなたはプロの戦略コンサルタントです。提供されたプロットを読み解き、スライド資料の構成を作成してください。
                ・1枚1テーマとし、タイトル（結論・主張）をつなげて読むだけでストーリーが成立するようにすること。
                ・各スライドの具体的なメッセージと要素を箇条書きで整理すること。
                """
                
                response_step1 = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=f"【プロット】\n{slide_plot}",
                    config=types.GenerateContentConfig(
                        system_instruction=step1_instruction,
                        temperature=0.2
                    )
                )
                story_text = response_step1.text

            # ==========================================
            # ステップ2：レイアウトの選定（デザイン担当）
            # ==========================================
            with st.spinner("ステップ2/3：AIがJinbaの62種類の型から最適なレイアウトを選定しています..."):
                step2_instruction = """
                あなたはプレゼン資料のデザイナーです。以下の【スライド構成】を読み、各スライドに最も適したレイアウト番号（layout_id）を割り当ててください。
                
                【Jinbaレイアウトカタログ（抜粋）】
                1: 通常の箇条書き
                11: 左右の比較表（メリット・デメリットなど）
                23: スケジュール・ロードマップ
                31: 3つのポイント（並列）
                45: 課題と解決策の対比
                ※迷った場合や該当するものがない場合は「1」を選択すること。
                """
                
                response_step2 = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=f"【スライド構成】\n{story_text}",
                    config=types.GenerateContentConfig(
                        system_instruction=step2_instruction,
                        temperature=0.2
                    )
                )
                layout_text = response_step2.text

            # ==========================================
            # ステップ3：JSON形式への清書（プログラマー担当）
            # ==========================================
            with st.spinner("ステップ3/3：AIがシステム用の最終データ（JSON）に変換しています..."):
                step3_instruction = """
                あなたはデータエンジニアです。これまでの【構成とレイアウト情報】をもとに、システムが読み込める完全なJSON配列を出力してください。
                出力形式は必ず以下のスキーマに従うこと。Markdownブロックは不要です。
                [
                  {
                    "layout_id": 1,
                    "title": "スライドのタイトル（主張）",
                    "content": ["ポイント1", "ポイント2", "ポイント3"]
                  }
                ]
                """
                
                response_step3 = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=f"【構成とレイアウト情報】\n{layout_text}",
                    config=types.GenerateContentConfig(
                        system_instruction=step3_instruction,
                        temperature=0.1,
                        response_mime_type="application/json"
                    )
                )
                
                # JSONの抽出とパース
                json_str = re.search(r'\[.*\]', response_step3.text, re.DOTALL)
                slide_data = json.loads(json_str.group(0)) if json_str else json.loads
