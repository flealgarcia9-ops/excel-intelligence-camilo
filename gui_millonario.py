#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import random
import tkinter as tk
from tkinter import messagebox

DATA_PATH = os.path.join(os.path.dirname(__file__), 'data', 'questions.json')


def load_questions(path: str = DATA_PATH):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_questions(questions, path: str = DATA_PATH):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(questions, f, ensure_ascii=False, indent=2)


class MillionaireGUI:
    def __init__(self, master, questions):
        self.master = master
        self.questions = questions
        self.total = min(len(questions), 10)
        self.ladder = [100, 200, 300, 500, 1000, 2000, 4000, 8000, 16000, 32000]
        self.index = 0
        self.lifeline_used = False

        self.question_var = tk.StringVar()
        self.options_vars = [tk.StringVar() for _ in range(4)]
        self.option_buttons = []  # store button refs for enable/disable during lifeline

        self._build_ui()
        self.show_question()

    def _build_ui(self):
        self.master.title("Quien quiere ser millonario - GUI")
        tk.Label(self.master, text="Premio:").grid(row=0, column=0, sticky='w')
        self.prize_label = tk.Label(self.master, text="0$")
        self.prize_label.grid(row=0, column=1, sticky='w')

        self.q_label = tk.Label(self.master, textvariable=self.question_var, wraplength=480, justify='left')
        self.q_label.grid(row=1, column=0, columnspan=2, sticky='w', pady=10)

        for i in range(4):
            btn = tk.Button(self.master, textvariable=self.options_vars[i], width=50, anchor='w', command=lambda idx=i: self.choose(idx))
            btn.grid(row=2+i, column=0, columnspan=2, padx=10, pady=5, sticky='w')
            self.option_buttons.append(btn)

        self.fifty_btn = tk.Button(self.master, text="50/50", command=self.use_5050)
        self.fifty_btn.grid(row=6, column=0, pady=10, sticky='w')
        self.phone_btn = tk.Button(self.master, text="Phone a Friend", command=self.use_phone_friend)
        self.phone_btn.grid(row=6, column=1, pady=10, sticky='w')
        self.status = tk.Label(self.master, text="")
        self.status.grid(row=7, column=0, sticky='w')

    def show_question(self):
        if self.index >= self.total:
            self.end_game()
            return
        q = self.questions[self.index]
        self.question_var.set(f"{self.index+1}. {q['question']}")
        for i in range(4):
            self.options_vars[i].set(q['options'][i])
            # ensure buttons are enabled for a new question unless lifeline was used previously
            self.option_buttons[i].config(state='normal')
        self.current_q = q
        self.visible_indices = [0,1,2,3]
        self.update_prize()

    def update_prize(self):
        self.prize_label.config(text=f"{self.ladder[self.index]}$")

    def end_game(self):
        messagebox.showinfo("Fin", f"Has terminado. Premio: {self.ladder[self.index-1] if self.index>0 else 0}$.")
        self.master.quit()

    def choose(self, idx):
        if idx not in self.visible_indices:
            return
        correct = int(self.current_q['answer'])
        if idx == correct:
            self.status.config(text="Correcto!")
            self.index += 1
            self.show_question()
        else:
            self.status.config(text="Incorrecto")
            self.end_game()

    def use_5050(self):
        if self.lifeline_used:
            messagebox.showinfo("Info", "Ya has usado 50/50 en esta ronda.")
            return
        self.lifeline_used = True
        correct = int(self.current_q['answer'])
        other = [i for i in range(4) if i != correct]
        import random as _random
        other_idx = _random.choice(other)
        self.visible_indices = sorted([correct, other_idx])
        # Disable non-visible options and clear their text
        for i, btn in enumerate(self.option_buttons):
            if i in self.visible_indices:
                btn.config(state='normal')
                self.options_vars[i].set(self.current_q['options'][i])
            else:
                btn.config(state='disabled')
                self.options_vars[i].set("")

    def use_phone_friend(self):
        if self.lifeline_used:
            messagebox.showinfo("Info", "Ya has usado una lifeline en esta ronda.")
            return
        self.lifeline_used = True
        q = self.current_q
        correct = int(q['answer'])
        other = [i for i in range(4) if i != correct]
        import random as _random
        other_idx = _random.choice(other)
        suggest = correct if _random.random() < 0.7 else other_idx
        suggestion = f"Phone a Friend sugiere: {chr(ord('A')+suggest)} - {q['options'][suggest]}"
        self.status.config(text=suggestion)


def main():
    questions = load_questions()
    root = tk.Tk()
    app = MillionaireGUI(root, questions)
    root.mainloop()


if __name__ == '__main__':
    main()
