import tkinter as tk
from tkinter import messagebox
import pandas as pd
import datetime
import random


class LanguageProcessingTestSystem:
    def __init__(self, root, window_size=(800, 600), font_size=16, font_family="Microsoft JhengHei", stage_order=["formal", "reward", "penalty", "reward_penalty"]):
        self.root = root
        self.root.title("詞彙判斷試驗系統")
        self.root.attributes("-fullscreen", True)  # 設置全屏顯示
        self.root.bind("<Escape>", self.exit_fullscreen)  # 綁定 Escape 鍵退出全屏

        self.participant_name = ""
        self.group = ""
        self.accuracy_threshold = 0.8
        self.stage_order = stage_order
        self.font = (font_family, font_size)  # 使用指定字體

        self.words = {
            "a": ["豆腐", "香菇", "菠菜"],  # 真詞
            "l": ["一一", "二二", "三三", "四四"],  # 假詞
            "space": {"雞蛋": 1, "雞心": 5, "雞胗": 7},  # PM target
        }
        self.true_words = set(self.words["a"])
        self.false_words = set(self.words["l"])
        self.pm_targets = self.words["space"]

        self.correct_answers = 0
        self.current_question_count = 0  # 當前已回答的問題數
        self.timeout_id = None
        self.current_stage_index = 0
        self.current_balance = 0  # 初始化金額

        # 計數器
        self.true_word_count = 0
        self.false_word_count = 0
        self.pm_target_count = 0
        self.true_word_correct = 0
        self.false_word_correct = 0
        self.pm_target_correct = 0

        # 答對率
        self.true_word_accuracy = 0
        self.false_word_accuracy = 0
        self.pm_target_accuracy = 0

        # 單詞清單，按指定順序排好
        self.word_list = self.create_word_list()

        # GUI設置
        self.setup_gui()

    def create_word_list(self):
        word_list = [None] * (len(self.true_words) + len(self.false_words) + len(self.pm_targets))

        # 將 PM target 移到指定的位置
        for target, pos in self.pm_targets.items():
            word_list[pos - 1] = (target, "space")

        # 將其餘的詞隨機填入空位
        remaining_words = []
        for word in self.true_words:
            remaining_words.append((word, "a"))
        for word in self.false_words:
            remaining_words.append((word, "l"))

        random.shuffle(remaining_words)

        # 將未指定順序的詞彙填充到空位
        for i in range(len(word_list)):
            if word_list[i] is None and remaining_words:
                word_list[i] = remaining_words.pop(0)

        return word_list

    def exit_fullscreen(self, event=None):
        """退出全屏"""
        self.root.attributes("-fullscreen", False)

    def setup_gui(self):
        """設置初始界面"""
        self.root.configure(bg="black")  # 設置背景為黑色

        self.instructions_label = tk.Label(
            self.root, text="歡迎來到詞彙判斷試驗系統", font=self.font, fg="white", bg="black"
        )
        self.instructions_label.pack()

        self.name_label = tk.Label(self.root, text="參與者姓名：", font=self.font, fg="white", bg="black")
        self.name_label.pack()
        self.name_entry = tk.Entry(self.root, font=self.font, fg="white", bg="black", insertbackground="white")
        self.name_entry.pack()

        self.group_label = tk.Label(self.root, text="組別：", font=self.font, fg="white", bg="black")
        self.group_label.pack()
        self.group_entry = tk.Entry(self.root, font=self.font, fg="white", bg="black", insertbackground="white")
        self.group_entry.pack()

        self.start_button = tk.Button(
            self.root, text="開始", command=self.start_experiment, font=self.font, fg="white", bg="black"
        )
        self.start_button.pack()

        # 顯示金額的Label
        self.balance_label = tk.Label(self.root, text="", font=self.font, fg="white", bg="black")
        self.balance_label.pack(anchor="nw", padx=10, pady=10)

    def update_balance_label(self):
        """更新金額顯示"""
        self.balance_label.config(text=f"金額: {self.current_balance} 元")
        self.balance_label.pack(anchor="nw", padx=10, pady=10)  # 確保顯示在左上角

    def start_experiment(self):
        """開始實驗"""
        self.participant_name = self.name_entry.get()
        self.group = self.group_entry.get()

        if not self.participant_name or not self.group:
            messagebox.showerror("錯誤", "請輸入姓名和組別。")
            return

        self.run_practice_instructions()

    def run_practice_instructions(self):
        """顯示練習指導語"""
        print(f'run_practice_instructions')
        # 清除所有現有的窗口內容
        for widget in self.root.winfo_children():
            widget.pack_forget()

        # 顯示前導詞
        self.instructions_label = tk.Label(
            self.root,
            text="請判斷螢幕上的詞彙是否為真實存在的詞彙，\n"
                 "真實的詞彙請按a，非真詞彙請按l，\n"
                 "然而，每當詞彙注音含有「歌、機」時，\n"
                 "請您按下空白建。\n"
                 "若您了解此實驗的程序請按enter鍵開始。",
            font=self.font,
            fg="white", bg="black"
        )
        self.instructions_label.pack(expand=True)

        # 綁定Enter鍵事件以開始練習
        self.root.bind("<Return>", self.show_any_key_screen)

    def show_any_key_screen(self, event):
        """顯示按任意鍵開始屏幕"""
        self.root.unbind("<Return>")
        for widget in self.root.winfo_children():
            widget.pack_forget()

        self.instructions_label = tk.Label(
            self.root,
            text="按任意鍵開始",
            font=self.font,
            fg="white", bg="black"
        )
        self.instructions_label.pack(expand=True)

        self.root.bind("<Key>", self.start_practice)

    def start_practice(self, event):
        """開始練習"""
        self.root.unbind("<Key>")
        self.run_practice()

    def run_practice(self):
        """運行練習"""
        self.word_list = self.create_word_list()  # 重置詞彙列表
        self.show_black_screen_before_next_word(stage="practice")

    def show_black_screen_before_next_word(self, stage):
        """顯示全黑屏幕500ms，然後顯示下一個單詞"""
        self.show_black_screen()
        self.root.after(500, lambda: self.show_next_word(stage))

    def show_black_screen(self):
        """顯示全黑屏幕"""
        for widget in self.root.winfo_children():
            widget.pack_forget()

        self.root.configure(bg="black")
        self.root.update()  # 确保界面更新以显示黑屏

    def show_next_word(self, stage):
        """顯示下一個單詞"""
        for widget in self.root.winfo_children():
            widget.pack_forget()

        if stage == "reward" or stage == "penalty" or stage == "reward_penalty":
            self.update_balance_label()  # 更新金額顯示

        if self.word_list:
            self.current_word, self.current_key = self.word_list.pop(0)
            self.instructions_label.config(text=self.current_word, font=self.font, fg="white", bg="black")
            self.instructions_label.pack(expand=True)
            self.root.bind("<Key>", lambda event: self.check_answer(event, stage))
            self.timeout_id = self.root.after(3000, lambda: self.check_answer_timeout(stage))
        else:
            self.end_stage(stage)

    def check_answer(self, event, stage):
        """檢查答案"""
        if self.timeout_id is not None:
            self.root.after_cancel(self.timeout_id)
            self.timeout_id = None

        self.root.unbind("<Key>")
        key = event.keysym.lower()

        if self.current_word in self.true_words:
            self.true_word_count += 1
            if key == "a":
                self.true_word_correct += 1
            self.true_word_accuracy = self.true_word_correct / self.true_word_count
            print(f'self.true_word_count: {self.true_word_count}')
            print(f'self.true_word_correct: {self.true_word_correct}')
            print(f'True word accuracy: {self.true_word_accuracy:.2%}')
        elif self.current_word in self.false_words:
            self.false_word_count += 1
            if key == "l":
                self.false_word_correct += 1
            self.false_word_accuracy = self.false_word_correct / self.false_word_count
            print(f'self.false_word_count: {self.false_word_count}')
            print(f'self.false_word_correct: {self.false_word_correct}')
            print(f'False word accuracy: {self.false_word_accuracy:.2%}')
        elif self.current_word in self.pm_targets:
            self.pm_target_count += 1
            if key == "space":
                self.pm_target_correct += 1
                if stage == "reward" or stage == "reward_penalty":
                    self.reward_user()  # 在獎勵或獎懲階段獎勵用戶
                    return  # 提前退出，防止继续执行
            else:
                if stage == "penalty" or stage == "reward_penalty":
                    self.penalize_user()  # 在懲罰或獎懲階段處罰用戶
                    return  # 提前退出，防止继续执行

        self.pm_target_accuracy = self.pm_target_correct / self.pm_target_count
        print(f'self.pm_target_count: {self.pm_target_count}')
        print(f'self.pm_target_correct: {self.pm_target_correct}')
        print(f'PM target accuracy: {self.pm_target_accuracy:.2%}')

        self.show_black_screen_before_next_word(stage)

    def reward_user(self):
        """獎勵用戶"""
        self.current_balance += 10
        self.show_reward_message()

    def show_reward_message(self):
        """顯示獎勵信息"""
        self.show_black_screen()
        self.instructions_label = tk.Label(
            self.root,
            text="獲得十元",
            font=self.font,
            fg="white", bg="black"
        )
        self.instructions_label.pack(expand=True)
        self.root.update()
        self.root.after(1500, lambda: self.update_balance_and_continue(stage="reward_penalty"))

    def penalize_user(self):
        """懲罰用戶"""
        self.current_balance -= 10
        self.show_penalty_message()

    def show_penalty_message(self):
        """顯示懲罰信息"""
        self.show_black_screen()
        self.instructions_label = tk.Label(
            self.root,
            text="扣除十元",
            font=self.font,
            fg="white", bg="black"
        )
        self.instructions_label.pack(expand=True)
        self.root.update()
        self.root.after(1500, lambda: self.update_balance_and_continue(stage="reward_penalty"))

    def update_balance_and_continue(self, stage):
        """更新金額並繼續"""
        self.update_balance_label()
        self.show_black_screen_before_next_word(stage)

    def check_answer_timeout(self, stage):
        """超時檢查答案"""
        if self.timeout_id is not None:
            self.timeout_id = None
            self.root.unbind("<Key>")

        if self.current_word in self.true_words:
            self.true_word_count += 1
            self.true_word_accuracy = self.true_word_correct / self.true_word_count
            print(f'self.true_word_count: {self.true_word_count}')
            print(f'self.true_word_correct: {self.true_word_correct}')
            print(f'True word accuracy: {self.true_word_accuracy:.2%}')
        elif self.current_word in self.false_words:
            self.false_word_count += 1
            self.false_word_accuracy = self.false_word_correct / self.false_word_count
            print(f'self.false_word_count: {self.false_word_count}')
            print(f'self.false_word_correct: {self.false_word_correct}')
            print(f'False word accuracy: {self.false_word_accuracy:.2%}')
        elif self.current_word in self.pm_targets:
            self.pm_target_count += 1
            if stage == "penalty" or stage == "reward_penalty":
                self.penalize_user()  # 只有在懲罰或獎懲階段才执行扣钱逻辑
                return  # 提前退出，防止继续执行

        self.pm_target_accuracy = self.pm_target_correct / self.pm_target_count
        print(f'self.pm_target_count: {self.pm_target_count}')
        print(f'self.pm_target_correct: {self.pm_target_correct}')
        print(f'PM target accuracy: {self.pm_target_accuracy:.2%}')

        self.show_black_screen_before_next_word(stage)

    def end_stage(self, stage):
        """結束階段"""
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
        """結束練習"""
        true_word_accuracy = self.true_word_correct / self.true_word_count if self.true_word_count else 0
        false_word_accuracy = self.false_word_correct / self.false_word_count if self.false_word_count else 0
        
        if true_word_accuracy < self.accuracy_threshold or false_word_accuracy < self.accuracy_threshold:
            self.reset_counters()
            self.run_practice_instructions()
        else:
            self.start_next_stage()

    def start_next_stage(self, event=None):
        """開始下一階段"""
        self.root.unbind("<Key>")
        self.run_main_experiment()

    def run_main_experiment(self):
        """運行主要實驗"""
        if self.current_stage_index < len(self.stage_order):
            stage = self.stage_order[self.current_stage_index]
            self.current_stage_index += 1
            if stage == "formal":
                print('formal')
                self.show_instructions("formal", 
                    "請判斷螢幕上的詞彙是否為真實存在的詞彙，\n"
                    "真實的詞彙請按a，非真詞彙請按l，\n"
                    "每當詞彙中注音含有「玻」與「的」時，\n"
                    "請您按下空白建。\n"
                    "若您了解此實驗的程序請按enter鍵開始。")
            elif stage == "reward":
                print('reward')
                self.show_instructions("reward", 
                    "請判斷螢幕上的詞彙是否為真實存在的詞彙，\n"
                    "真實的詞彙請按a，非真詞彙請按l，\n"
                    "每當詞彙中注音含有「波」與「特」時，\n"
                    "請您按下空白建。\n"
                    "每當您正確辨認出含有「波」與「特」的詞彙時，\n"
                    "會顯示您獲得10元，且會累計顯示於左上角，\n"
                    "若您了解此實驗的程序請按enter鍵開始。")
            elif stage == "penalty":
                print('penalty')
                self.show_instructions("penalty", 
                    "請判斷螢幕上的詞彙是否為真實存在的詞彙，\n"
                    "真實的詞彙請按a，非真詞彙請按l，\n"
                    "每當詞彙中注音含有「摸」與「呢」時，\n"
                    "請您按下空白建。\n"
                    "每當您未正確辨認出含有「摸」與「呢」的詞彙時，\n"
                    "會顯示您被扣除10元，且會累計顯示於左上角，\n"
                    "若您了解此實驗的程序請按enter鍵開始。")
            elif stage == "reward_penalty":
                print('reward_penalty')
                self.show_instructions("reward_penalty", 
                    "請判斷螢幕上的詞彙是否為真實存在的詞彙，\n"
                    "真實的詞彙請按a，非真詞彙請按l，\n"
                    "每當詞彙中注音含有「佛」與「了」時，\n"
                    "請您按下空白建。\n"
                    "每當您正確辨認出含有「佛」的詞彙時，\n"
                    "會顯示您獲得10元，且會累計顯示於左上角，\n"
                    "每當您未正確辨認出含有「了」的詞彙時，\n"
                    "會顯示您被扣除10元，且會累計顯示於左上角，\n"
                    "若您了解此實驗的程序請按enter鍵開始。")
        else:
            self.save_results()
            self.show_thank_you_message()

    def show_instructions(self, stage, instructions):
        """顯示每個階段的指導語"""
        self.instructions_label.config(text=instructions, font=self.font, fg="white", bg="black")
        self.instructions_label.pack(expand=True)
        self.root.bind("<Return>", lambda event: self.show_any_key_screen_next(event, stage))

    def show_any_key_screen_next(self, event, stage):
        """顯示按任意鍵開始屏幕"""
        self.root.unbind("<Return>")
        for widget in self.root.winfo_children():
            widget.pack_forget()

        self.instructions_label = tk.Label(
            self.root,
            text="按任意鍵開始",
            font=self.font,
            fg="white", bg="black"
        )
        self.instructions_label.pack(expand=True)

        self.root.bind("<Key>", lambda event: self.start_stage(event, stage))

    def start_stage(self, event, stage):
        """開始階段"""
        self.root.unbind("<Key>")
        self.word_list = self.create_word_list()  # 重置詞彙列表
        if stage == "formal":
            self.run_formal_stage()
        elif stage == "reward":
            self.run_reward_stage()
        elif stage == "penalty":
            self.run_penalty_stage()
        elif stage == "reward_penalty":
            self.run_reward_penalty_stage()

    def run_formal_stage(self):
        """運行正式測試階段"""
        self.reset_counters()
        self.show_black_screen_before_next_word(stage="formal")

    def end_formal_stage(self):
        """結束正式測試階段"""
        self.start_next_stage()

    def run_reward_stage(self):
        """運行獎勵階段"""
        self.reset_counters()
        self.current_balance = 200  # 初始金額
        self.update_balance_label()  # 顯示金額
        self.show_black_screen_before_next_word(stage="reward")

    def end_reward_stage(self):
        """結束獎勵階段"""
        self.clear_balance_label()  # 清除金額顯示
        self.start_next_stage()

    def run_penalty_stage(self):
        """運行懲罰階段"""
        self.reset_counters()
        self.current_balance = 200  # 初始金額
        self.update_balance_label()  # 顯示金額
        self.show_black_screen_before_next_word(stage="penalty")

    def end_penalty_stage(self):
        """結束懲罰階段"""
        self.clear_balance_label()  # 清除金額顯示
        self.start_next_stage()

    def run_reward_penalty_stage(self):
        """運行獎懲階段"""
        self.reset_counters()
        self.current_balance = 200  # 初始金額
        self.update_balance_label()  # 顯示金額
        self.show_black_screen_before_next_word(stage="reward_penalty")

    def end_reward_penalty_stage(self):
        """結束獎懲階段"""
        self.clear_balance_label()  # 清除金額顯示
        self.start_next_stage()

    def clear_balance_label(self):
        """清除金額顯示"""
        self.balance_label.config(text="")

    def reset_counters(self):
        """重置計數器"""
        self.correct_answers = 0
        self.current_question_count = 0
        self.true_word_count = 0
        self.false_word_count = 0
        self.pm_target_count = 0
        self.true_word_correct = 0
        self.false_word_correct = 0
        self.pm_target_correct = 0

    def save_results(self):
        """保存結果"""
        data = {
            "姓名": [self.participant_name],
            "組別": [self.group],
            "日期": [datetime.datetime.now().strftime("%Y-%m-%d")],
            "正確回答數": [self.correct_answers],
            "總回答數": [self.true_word_count + self.false_word_count + self.pm_target_count],
            "正確率": [self.correct_answers / (self.true_word_count + self.false_word_count + self.pm_target_count) * 100],
            "真詞正確率": [self.true_word_correct / self.true_word_count * 100 if self.true_word_count else 0],
            "假詞正確率": [self.false_word_correct / self.false_word_count * 100 if self.false_word_count else 0],
            "PM target正確率": [self.pm_target_correct / self.pm_target_count * 100 if self.pm_target_count else 0]
        }
        df = pd.DataFrame(data)
        df.to_excel(f"{self.group}_{self.participant_name}_results.xlsx", index=False)
        messagebox.showinfo("結果已保存", "結果已保存到Excel文件中。")

    def show_thank_you_message(self):
        """顯示銘謝詞並停留"""
        for widget in self.root.winfo_children():
            widget.pack_forget()

        self.instructions_label = tk.Label(
            self.root,
            text="感謝您的參與，測驗已完成。",
            font=self.font,
            fg="white", bg="black"
        )
        self.instructions_label.pack(expand=True)
        # 不綁定任何事件，停留在此屏幕

if __name__ == "__main__":
    root = tk.Tk()
    app = LanguageProcessingTestSystem(
        root,
        window_size=(1024, 768),
        font_size=32,
        font_family="Microsoft JhengHei",
        stage_order=["penalty", "reward", "reward_penalty", "formal", ],
    )
    root.mainloop()
