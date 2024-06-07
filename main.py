import tkinter as tk
from tkinter import messagebox
import pandas as pd
import datetime
import random


class LanguageProcessingTestSystem:
    def __init__(
        self,
        root,
        window_size=(800, 600),
        font_size=16,
        font_family="Microsoft JhengHei",
        stage_order=["formal", "reward", "penalty", "reward_penalty"],
        total_answers=5,  # 每個區段的總題數
    ):
        self.root = root
        self.root.title("詞彙判斷試驗系統")
        self.root.geometry(f"{window_size[0]}x{window_size[1]}")

        self.participant_name = ""
        self.group = ""
        self.accuracy_threshold = 0.8
        self.stage_order = stage_order
        self.font = (font_family, font_size)  # 使用指定字體

        self.words = {
            "a": ["豆腐", "香菇", "菠菜"],
            "b": ["一一", "二二", "三三", "四四"],
            "space": ["雞蛋"],
        }
        self.correct_answers = 0
        self.total_answers = total_answers  # 每個區段的總題數
        self.current_question_count = 0  # 當前已回答的問題數
        self.timeout_id = None
        self.current_stage_index = 0

        # GUI設置
        self.setup_gui()

    def setup_gui(self):
        self.instructions_label = tk.Label(
            self.root, text="歡迎來到詞彙判斷試驗系統", font=self.font
        )
        self.instructions_label.pack()

        self.name_label = tk.Label(self.root, text="參與者姓名：", font=self.font)
        self.name_label.pack()
        self.name_entry = tk.Entry(self.root, font=self.font)
        self.name_entry.pack()

        self.group_label = tk.Label(self.root, text="組別：", font=self.font)
        self.group_label.pack()
        self.group_entry = tk.Entry(self.root, font=self.font)
        self.group_entry.pack()

        self.start_button = tk.Button(
            self.root, text="開始", command=self.start_experiment, font=self.font
        )
        self.start_button.pack()

    def start_experiment(self):
        self.participant_name = self.name_entry.get()
        self.group = self.group_entry.get()

        if not self.participant_name or not self.group:
            messagebox.showerror("錯誤", "請輸入姓名和組別。")
            return

        self.run_practice_instructions()

    def run_practice_instructions(self):
        # 清除所有現有的窗口內容
        for widget in self.root.winfo_children():
            widget.pack_forget()

        # 顯示前導詞
        self.instructions_label = tk.Label(
            self.root,
            text="指示：您將看到一些單詞，需要進行識別。按任意鍵開始練習。",
            font=self.font,
        )
        self.instructions_label.pack(expand=True)

        # 綁定任意鍵事件以開始練習
        self.root.bind("<Key>", self.start_practice)

    def start_practice(self, event):
        self.root.unbind("<Key>")
        self.run_practice()

    def run_practice(self):
        self.show_random_word(stage="practice")

    def show_random_word(self, stage):
        self.current_word = random.choice(
            [word for key in self.words for word in self.words[key]]
        )
        self.instructions_label.config(text=self.current_word, font=self.font)
        self.root.bind("<Key>", lambda event: self.check_answer(event, stage))
        self.timeout_id = self.root.after(3000, lambda: self.check_answer_timeout(stage))

    def check_answer(self, event, stage):
        if self.timeout_id is not None:
            self.root.after_cancel(self.timeout_id)
            self.timeout_id = None

        self.root.unbind("<Key>")
        key = event.keysym.lower()
        correct_key = None

        for k, v in self.words.items():
            if self.current_word in v:
                correct_key = k
                break

        if key == correct_key:
            self.correct_answers += 1
            result_text = "正確！"
        else:
            result_text = "錯誤！"

        self.current_question_count += 1
        self.show_result(result_text, stage)

    def check_answer_timeout(self, stage):
        if self.timeout_id is not None:
            self.timeout_id = None
            self.root.unbind("<Key>")
            result_text = "錯誤！"
            self.current_question_count += 1
            self.show_result(result_text, stage)

    def show_result(self, result_text, stage):
        accuracy = (
            self.correct_answers / self.current_question_count * 100
        )
        self.instructions_label.config(
            text=f"{result_text}\n\n正確回答數：{self.correct_answers}\n總回答數：{self.current_question_count}\n正確率：{accuracy:.2f}%",
            font=self.font,
        )
        self.root.after(1000, lambda: self.show_next_word(stage))

    def show_next_word(self, stage):
        if self.current_question_count < self.total_answers:
            self.show_random_word(stage)
        else:
            if stage == "practice":
                self.end_practice()
            elif stage == "formal":
                self.end_formal_stage()
            elif stage == "reward":
                self.end_reward_stage()
            elif stage == "penalty":
                self.end_penalty_stage()
            elif stage == "reward_penalty":
                self.end_reward_penalty_stage()

    def end_practice(self):
        accuracy = self.correct_answers / self.total_answers
        if accuracy < self.accuracy_threshold:
            self.instructions_label.config(
                text=f"正確率不足80%，重新練習。\n\n最終結果：\n正確回答數：{self.correct_answers}\n總回答數：{self.total_answers}\n正確率：{accuracy * 100:.2f}%\n\n按任意鍵重新練習。",
                font=self.font,
            )
            self.root.bind("<Key>", self.start_practice_again)
        else:
            self.instructions_label.config(
                text=f"練習完成。\n\n最終結果：\n正確回答數：{self.correct_answers}\n總回答數：{self.total_answers}\n正確率：{accuracy * 100:.2f}%\n\n按任意鍵進入正式區。",
                font=self.font,
            )
            self.root.bind("<Key>", self.start_next_stage)

    def start_practice_again(self, event):
        self.root.unbind("<Key>")
        self.reset_counters()
        self.run_practice()

    def start_formal_stage(self, event):
        self.root.unbind("<Key>")
        self.run_formal_stage()

    def run_formal_stage(self):
        self.reset_counters()
        self.show_random_word(stage="formal")

    def end_formal_stage(self):
        accuracy = self.correct_answers / self.total_answers
        self.instructions_label.config(
            text=f"正式區完成。\n\n最終結果：\n正確回答數：{self.correct_answers}\n總回答數：{self.total_answers}\n正確率：{accuracy * 100:.2f}%\n\n按任意鍵進入下一區。",
            font=self.font,
        )
        self.root.bind("<Key>", self.start_next_stage)

    def end_reward_stage(self):
        self.instructions_label.config(
            text="獎勵區完成。\n\n按任意鍵進入下一區。",
            font=self.font,
        )
        self.root.bind("<Key>", self.start_next_stage)

    def end_penalty_stage(self):
        self.instructions_label.config(
            text="懲罰區完成。\n\n按任意鍵進入下一區。",
            font=self.font,
        )
        self.root.bind("<Key>", self.start_next_stage)

    def end_reward_penalty_stage(self):
        self.instructions_label.config(
            text="獎懲區完成。\n\n按任意鍵進入下一區。",
            font=self.font,
        )
        self.root.bind("<Key>", self.start_next_stage)

    def start_next_stage(self, event):
        self.root.unbind("<Key>")
        self.run_main_experiment()

    def run_main_experiment(self):
        if self.current_stage_index < len(self.stage_order):
            stage = self.stage_order[self.current_stage_index]
            self.current_stage_index += 1
            if stage == "formal":
                self.run_formal_stage()
            elif stage == "reward":
                self.run_reward_stage()
            elif stage == "penalty":
                self.run_penalty_stage()
            elif stage == "reward_penalty":
                self.run_reward_penalty_stage()
        else:
            self.instructions_label.config(text="測試完成。謝謝！", font=self.font)
            self.save_results()

    def run_reward_stage(self):
        self.reset_counters()
        self.instructions_label.config(text="獎勵區\n\n按任意鍵開始。", font=self.font)
        self.root.bind("<Key>", self.start_reward_stage)

    def start_reward_stage(self, event):
        self.root.unbind("<Key>")
        self.show_random_word(stage="reward")

    def run_penalty_stage(self):
        self.reset_counters()
        self.instructions_label.config(text="懲罰區\n\n按任意鍵開始。", font=self.font)
        self.root.bind("<Key>", self.start_penalty_stage)

    def start_penalty_stage(self, event):
        self.root.unbind("<Key>")
        self.show_random_word(stage="penalty")

    def run_reward_penalty_stage(self):
        self.reset_counters()
        self.instructions_label.config(text="獎懲區\n\n按任意鍵開始。", font=self.font)
        self.root.bind("<Key>", self.start_reward_penalty_stage)

    def start_reward_penalty_stage(self, event):
        self.root.unbind("<Key>")
        self.show_random_word(stage="reward_penalty")

    def reset_counters(self):
        self.correct_answers = 0
        self.current_question_count = 0

    def save_results(self):
        data = {
            "姓名": [self.participant_name],
            "組別": [self.group],
            "日期": [datetime.datetime.now().strftime("%Y-%m-%d")],
            # Add other columns as needed
        }
        df = pd.DataFrame(data)
        df.to_excel(f"{self.participant_name}_results.xlsx", index=False)
        messagebox.showinfo("結果已保存", "結果已保存到Excel文件中。")


if __name__ == "__main__":
    root = tk.Tk()
    app = LanguageProcessingTestSystem(
        root,
        window_size=(1024, 768),
        font_size=32,
        font_family="Microsoft JhengHei",
        stage_order=["reward_penalty", "reward", "penalty",  "formal",],
        total_answers=5,
    )
    root.mainloop()
