#!/usr/bin/env python3
"""Quien quiere ser millonario - CLI versión simple con lifeline 50/50.

Funciones auxiliares para cargar preguntas, añadir, listar y jugar.
"""
from __future__ import annotations

import json
import os
import random
from typing import List, Dict, Optional, Tuple

DATA_PATH = os.path.join(os.path.dirname(__file__), 'data', 'questions.json')


def load_questions(path: Optional[str] = None) -> List[Dict]:
    path = path or DATA_PATH
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_questions(questions: List[Dict], path: Optional[str] = None) -> None:
    path = path or DATA_PATH
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(questions, f, ensure_ascii=False, indent=2)


def add_question(questions: List[Dict], qtext: str, options: List[str], answer_idx: int) -> List[Dict]:
    new_id = max((q.get('id', 0) for q in questions), default=0) + 1
    questions.append({"id": new_id, "question": qtext, "options": options, "answer": int(answer_idx)})
    return questions


def list_questions(questions: List[Dict]) -> str:
    lines = []
    for q in questions:
        idx = q.get('id', '?')
        lines.append(f"{idx}: {q['question']} (A:{q['options'][0]} B:{q['options'][1]} C:{q['options'][2]} D:{q['options'][3]})")
    return "\n".join(lines)


def delete_question_by_id(questions: List[Dict], qid: int) -> List[Dict]:
    return [q for q in questions if int(q.get('id')) != int(qid)]


def get_5050_indices(question: Dict, rng: Optional[random.Random] = None) -> Tuple[int, int]:
    rng = rng or random
    correct = int(question['answer'])
    options = [i for i in range(4) if i != correct]
    other = rng.choice(options)
    return correct, other


def print_help():
    help_text = (
        "Comandos:\n"
        "  play            Jugar una ronda con las preguntas guardadas\n"
        "  add             Añadir pregunta interactiva\n"
        "  list            Listar preguntas guardadas\n"
        "  delete <id>     Eliminar pregunta por id\n"
        "  help            Mostrar esta ayuda\n"
    )
    print(help_text)


def add_question_interactive(path: Optional[str] = None) -> None:
    questions = load_questions(path)
    qtext = input("Pregunta: ")
    opts = []
    for i in range(4):
        o = input(f"Opción {chr(ord('A')+i)}: ")
        opts.append(o)
    ans = None
    while ans is None:
        a = input("Respuesta correcta (A-D): ")
        if a and a.upper() in 'ABCD':
            ans = ord(a.upper()) - ord('A')
        else:
            print("Entrada inválida. Elige A-D.")
    questions = add_question(questions, qtext, opts, int(ans))
    save_questions(questions, path)
    print("Pregunta añadida con éxito.")


def play_game(questions: Optional[List[Dict]] = None, max_questions: int = 10, rng: Optional[random.Random] = None) -> None:
    rng = rng or random
    questions = questions if questions is not None else load_questions()
    if not questions:
        print("No hay preguntas disponibles.")
        return
    total = min(len(questions), max_questions)
    ladder = [100, 200, 300, 500, 1000, 2000, 4000, 8000, 16000, 32000]
    score = 0
    lifeline_used = False

    for i in range(total):
        q = questions[i]
        print(f"Pregunta {i+1} - {ladder[i]}$")
        print(f"{q['question']}")
        for idx, opt in enumerate(q['options']):
            print(f"{chr(ord('A')+idx)}) {opt}")

        visible = [0,1,2,3]
        answer = None
        while True:
            user = input("Tu respuesta (A-D) o '50' para 50/50 o 'phone' para Phone a Friend: ").strip().lower()
            if user == '50':
                if lifeline_used:
                    print("Ya has usado la lifeline 50/50 en esta ronda.")
                    continue
                lifeline_used = True
                cidx = int(q['answer'])
                other = [x for x in range(4) if x != cidx][0]
                # choose random incorrect to display with correct one
                other = random.choice([x for x in range(4) if x != cidx])
                visible = sorted([cidx, other])
                print("50/50: opciones disponibles:")
                for idx in visible:
                    print(f"{chr(ord('A')+idx)}) {q['options'][idx]}")
                continue
            if user == 'phone':
                if lifeline_used:
                    print("Ya has usado una lifeline en esta ronda.")
                    continue
                lifeline_used = True
                correct = int(q['answer'])
                other = [i for i in range(4) if i != correct]
                other_idx = random.choice(other)
                suggest = correct if random.random() < 0.7 else other_idx
                print(f"Phone a Friend sugiere: {chr(ord('A')+suggest)} - {q['options'][suggest]}")
                continue
            if user and user[0] in 'abcd':
                answer = ord(user[0].upper()) - ord('A')
                if answer not in visible:
                    print("Opción no válida en este momento.")
                    continue
                break
            print("Entrada no válida. Usa A-D o 50.")

        if answer == int(q['answer']):
            score = ladder[i]
            print(f"Correcto! Sumas {score}$")
            if i == total - 1:
                print("¡Has ganado el premio máximo!")
        else:
            print("Incorrecto. Fin del juego.")
            break

    print(f"Juego terminado. Premio obtenido: {score}$")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Quien quiere ser millonario - CLI")
    sub = parser.add_subparsers(dest='cmd')
    sub.add_parser('play', help='Jugar ronda')
    sub.add_parser('add', help='Añadir pregunta interactiva')
    sub.add_parser('list', help='Listar preguntas')
    del_parser = sub.add_parser('delete', help='Eliminar pregunta por id')
    del_parser.add_argument('id', type=int, help='ID de la pregunta')
    sub.add_parser('help', help='Mostrar ayuda')
    args = parser.parse_args()

    qs = load_questions()
    if args.cmd == 'play':
        play_game(questions=qs)
    elif args.cmd == 'add':
        add_question_interactive()
    elif args.cmd == 'list':
        print(list_questions(qs))
    elif args.cmd == 'delete':
        new = delete_question_by_id(qs, args.id)
        save_questions(new)
        print(f"Pregunta {args.id} eliminada (si existía).")
    else:
        print_help()


if __name__ == '__main__':
    main()
