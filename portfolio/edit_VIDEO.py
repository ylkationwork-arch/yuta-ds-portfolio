#AIzaSyAkoK0cmge8SmlF8ugLCmieiX6LkzqsKSc
import os
import time
import json
import threading
import requests
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import google.generativeai as genai
from moviepy.editor import VideoFileClip, AudioFileClip, CompositeAudioClip, concatenate_videoclips, vfx

# --- 設定 ---
VOICEVOX_URL = "http://127.0.0.1:50021"
SPEAKER_ID = 3  # ずんだもん（ノーマル）

# --- ヘルパー関数: 時間変換 ---
def parse_time(time_str):
    """ MM:SS または SS 形式の文字列を秒(float)に変換 """
    try:
        if ":" in str(time_str):
            parts = str(time_str).split(":")
            return int(parts[0]) * 60 + float(parts[1])
        return float(time_str)
    except Exception:
        return 0.0

# --- クラス: 動画編集アプリ ---
class AIEditorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gemini x Zundamon 自動編集エディタ")
        self.geometry("700x650")
        ctk.set_appearance_mode("Dark")
        
        # 変数
        self.video_path = tk.StringVar()
        self.api_key = tk.StringVar()
        
        # レイアウト作成
        self.create_widgets()

    def create_widgets(self):
        # 1. API Key入力
        self.frame_api = ctk.CTkFrame(self)
        self.frame_api.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(self.frame_api, text="Gemini API Key:").pack(side="left", padx=10)
        self.entry_api = ctk.CTkEntry(self.frame_api, textvariable=self.api_key, show="*", width=400)
        self.entry_api.pack(side="left", padx=10)

        # 2. ファイル選択
        self.frame_file = ctk.CTkFrame(self)
        self.frame_file.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkButton(self.frame_file, text="動画ファイルを選択 (MP4)", command=self.select_file).pack(side="left", padx=10)
        self.lbl_file = ctk.CTkLabel(self.frame_file, textvariable=self.video_path)
        self.lbl_file.pack(side="left", padx=10)

        # 3. 編集方針（プロンプト）
        ctk.CTkLabel(self, text="編集方針 (Geminiへの指示):").pack(pady=(10, 0), padx=20, anchor="w")
        self.txt_prompt = ctk.CTkTextbox(self, height=100)
        self.txt_prompt.pack(pady=5, padx=20, fill="x")
        self.txt_prompt.insert("1.0", "この動画から面白いシーンや重要なシーンを3つ程度選んで切り抜いてください。ずんだもんが視聴者に話しかけるような実況スタイルで解説テキストを作成してください。")

        # 4. 実行ボタン
        self.btn_run = ctk.CTkButton(self, text="自動編集を開始する", command=self.start_processing_thread, fg_color="green", height=40)
        self.btn_run.pack(pady=20, padx=20, fill="x")

        # 5. ログ表示
        ctk.CTkLabel(self, text="実行ログ:").pack(pady=(10, 0), padx=20, anchor="w")
        self.txt_log = ctk.CTkTextbox(self, height=200)
        self.txt_log.pack(pady=5, padx=20, fill="both", expand=True)

    def log(self, message):
        self.txt_log.insert("end", message + "\n")
        self.txt_log.see("end")

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Video Files", "*.mp4 *.mov *.avi")])
        if file_path:
            self.video_path.set(file_path)

    def start_processing_thread(self):
        if not self.api_key.get() or not self.video_path.get():
            messagebox.showerror("エラー", "APIキーと動画ファイルを設定してください")
            return
        
        # UIフリーズ防止のため別スレッドで実行
        self.btn_run.configure(state="disabled")
        threading.Thread(target=self.run_process, daemon=True).start()

    # --- メイン処理ロジック ---
    def run_process(self):
        try:
            api_key = self.api_key.get()
            video_path = self.video_path.get()
            user_prompt = self.txt_prompt.get("1.0", "end").strip()
            
            genai.configure(api_key=api_key)

            # 1. 動画のアップロード
            self.log("--- 処理開始 ---")
            self.log("1. Geminiへ動画をアップロード中...")
            video_file = genai.upload_file(path=video_path)
            
            # 処理完了待ち
            while video_file.state.name == "PROCESSING":
                time.sleep(2)
                video_file = genai.get_file(video_file.name)
            
            if video_file.state.name == "FAILED":
                raise Exception("動画の処理に失敗しました")

            self.log("動画の解析準備完了。AIによる編集プラン作成中...")

            # 2. Geminiによる構成案作成
            model = genai.GenerativeModel(model_name="gemini-2.5-flash")
            
            # JSON形式を強制するためのシステムプロンプト
            system_instruction = """
            あなたは動画編集者です。提供された動画とユーザーの指示に基づいて、動画を編集するためのプランをJSON形式で出力してください。
            出力は以下のキーを持つオブジェクトのリスト（配列）のみにしてください。Markdown記法は不要です。
            [
              {
                "start": "MM:SS",
                "end": "MM:SS",
                "script": "ずんだもんが話すセリフ"
              }
            ]
            ・startとendは動画の切り抜き範囲です。
            ・scriptはVOICEVOXで読み上げる内容です。
            """
            
            response = model.generate_content(
                [video_file, system_instruction, f"ユーザーの指示: {user_prompt}"],
                generation_config={"response_mime_type": "application/json"}
            )
            
            edit_plan = json.loads(response.text)
            self.log(f"編集プラン取得: {len(edit_plan)}個のカットを作成します")

            # 3. VOICEVOX音声生成と動画編集
            final_clips = []
            original_clip = VideoFileClip(video_path)

            for i, item in enumerate(edit_plan):
                start_t = parse_time(item['start'])
                end_t = parse_time(item['end'])
                script = item['script']
                
                self.log(f"カット{i+1}: {item['start']}~{item['end']} 「{script}」")

                # A. 音声生成 (VOICEVOX)
                audio_path = self.generate_voicevox_audio(script, f"temp_audio_{i}.wav")
                
                # B. 動画カット
                # 時間指定が範囲外でないかチェック
                if start_t >= original_clip.duration: continue
                if end_t > original_clip.duration: end_t = original_clip.duration
                
                sub_clip = original_clip.subclip(start_t, end_t)
                
                # C. 音声合成 (元の音声を小さく + ずんだもん)
                voice_audio = AudioFileClip(audio_path)
                
                # 元動画の音量を下げる (20%にする)
                bgm_audio = sub_clip.audio.volumex(0.2) if sub_clip.audio else None
                
                # 動画と音声の長さを調整
                # もし実況が動画より長ければ、動画の最後のフレームを静止画として延長する
                if voice_audio.duration > sub_clip.duration:
                    duration_diff = voice_audio.duration - sub_clip.duration
                    # 最後のフレームを取得して静止画クリップ作成
                    last_frame = sub_clip.to_ImageClip(duration=duration_diff)
                    last_frame.fps = 24
                    # 結合
                    sub_clip = concatenate_videoclips([sub_clip, last_frame])
                    # 元音声を延長（あるいは静止画部分は無音）
                    # 簡易化のため、BGMは動画の長さ分だけにする(静止画部分は無音orループ)
                    # ここではシンプルに、実況の長さに合わせる
                    final_duration = voice_audio.duration
                else:
                    final_duration = sub_clip.duration
                
                # 音声のミックス
                if bgm_audio:
                    # BGMが短い場合はそのまま、長い場合はカット
                    bgm_audio = bgm_audio.set_duration(min(bgm_audio.duration, final_duration))
                    final_audio = CompositeAudioClip([bgm_audio, voice_audio])
                else:
                    final_audio = voice_audio

                sub_clip = sub_clip.set_audio(final_audio)
                sub_clip = sub_clip.set_duration(final_duration)
                
                final_clips.append(sub_clip)

            # 4. 全クリップ結合と書き出し
            self.log("全クリップを結合して書き出し中...")
            final_video = concatenate_videoclips(final_clips)
            
            output_filename = "output_zundamon_edit.mp4"
            final_video.write_videofile(
                output_filename, 
                codec="libx264", 
                audio_codec="aac",
                fps=24,
                logger=None # プログレスバーをコンソールに出さない
            )
            
            self.log(f"完了！ファイルが出力されました: {output_filename}")
            messagebox.showinfo("成功", "動画の編集が完了しました！")

            # 後片付け (一時ファイルの削除など)
            original_clip.close()
            for i in range(len(edit_plan)):
                if os.path.exists(f"temp_audio_{i}.wav"):
                    os.remove(f"temp_audio_{i}.wav")

        except Exception as e:
            self.log(f"エラー発生: {str(e)}")
            messagebox.showerror("エラー", f"処理中にエラーが発生しました\n{str(e)}")
        
        finally:
            self.btn_run.configure(state="normal")

    def generate_voicevox_audio(self, text, output_path):
        """ VOICEVOX APIを叩いてwavファイルを保存する """
        # Audio Query
        params = {'text': text, 'speaker': SPEAKER_ID}
        query_res = requests.post(f"{VOICEVOX_URL}/audio_query", params=params)
        query_json = query_res.json()
        
        # Synthesis
        synthesis_res = requests.post(
            f"{VOICEVOX_URL}/synthesis", 
            headers={"Content-Type": "application/json"},
            params={'speaker': SPEAKER_ID},
            data=json.dumps(query_json)
        )
        
        with open(output_path, "wb") as f:
            f.write(synthesis_res.content)
        
        return output_path

if __name__ == "__main__":
    app = AIEditorApp()
    app.mainloop()