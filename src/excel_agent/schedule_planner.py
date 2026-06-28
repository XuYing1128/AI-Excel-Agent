"""Deterministic greedy scheduler for exam-style assignments.

Given a list of (student_id, subject_code) registrations, a set of rooms with
capacity, and a set of time slots, assign each registration to one (slot, room)
such that no student is in two rooms at the same time and rooms never exceed
their seat count.

Simple, dependency-free, and prompt-driven: room names, time slots and capacity
are parsed from plain-language descriptions when present and otherwise default
to the configuration the user gave us in the conversation (3 rooms × 65 seats,
two days × three slots).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class TimeSlot:
    date_label: str   # e.g. "第一天"
    start: str        # "08:00"
    end: str          # "10:00"

    def label(self) -> str:
        return f"{self.date_label} {self.start}-{self.end}"


DEFAULT_ROOMS = ["二教E504", "二教E505", "二教E506"]
DEFAULT_SEATS_PER_ROOM = 65
DEFAULT_SLOTS: list[TimeSlot] = [
    TimeSlot("第一天", "08:00", "10:00"),
    TimeSlot("第一天", "10:30", "12:30"),
    TimeSlot("第一天", "14:30", "16:30"),
    TimeSlot("第二天", "08:00", "10:00"),
    TimeSlot("第二天", "10:30", "12:30"),
    TimeSlot("第二天", "14:30", "16:30"),
]


@dataclass
class Assignment:
    room: str
    date_label: str
    start: str
    end: str


def parse_rooms(prompt: str) -> list[str]:
    """Pick out concrete room names from the user's description."""

    rooms: list[str] = []
    for match in re.finditer(r"[一二三四五六七八九十]?教\s*[A-Za-z]?\d{2,4}", prompt or ""):
        name = match.group(0).replace(" ", "")
        if name not in rooms:
            rooms.append(name)
    return rooms or list(DEFAULT_ROOMS)


def parse_capacity(prompt: str) -> int:
    """Pick out seats-per-room ("一个教室65个位置")."""

    match = re.search(r"(\d{2,4})\s*(?:个)?(?:位置|座位|位子|人)", prompt or "")
    if match:
        try:
            return max(1, int(match.group(1)))
        except ValueError:
            pass
    return DEFAULT_SEATS_PER_ROOM


def parse_slots(prompt: str, day_count: int = 2) -> list[TimeSlot]:
    """Pick HH:MM-HH:MM windows from the prompt, mirror across the day count."""

    text = str(prompt or "").replace("：", ":").replace("．", ".")
    pairs: list[tuple[str, str]] = []
    pattern = re.compile(
        r"(\d{1,2})\s*(?:点|:)\s*(\d{0,2})\s*(?:到|-|—|至|~)\s*"
        r"(\d{1,2})\s*(?:点|:)\s*(\d{0,2})"
    )
    for match in pattern.finditer(text):
        s_h, s_m, e_h, e_m = match.groups()
        start = f"{int(s_h):02d}:{int(s_m or 0):02d}"
        end = f"{int(e_h):02d}:{int(e_m or 0):02d}"
        pair = (start, end)
        if pair not in pairs:
            pairs.append(pair)
    if not pairs:
        return list(DEFAULT_SLOTS)
    slots: list[TimeSlot] = []
    day_labels = [f"第{c}天" for c in "一二三四五六"][:max(1, day_count)]
    for day in day_labels:
        for start, end in pairs:
            slots.append(TimeSlot(day, start, end))
    return slots or list(DEFAULT_SLOTS)


def parse_day_count(prompt: str) -> int:
    match = re.search(r"([两二三四五六两])天考完|(\d+)\s*天", str(prompt or ""))
    if not match:
        return 2
    word = match.group(1) or match.group(2) or ""
    table = {"两": 2, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6}
    if word in table:
        return table[word]
    try:
        return max(1, int(word))
    except ValueError:
        return 2


def schedule_registrations(
    registrations: list[dict[str, Any]],
    *,
    student_key: str = "准考证号",
    rooms: list[str] | None = None,
    seats_per_room: int = DEFAULT_SEATS_PER_ROOM,
    slots: list[TimeSlot] | None = None,
) -> list[Assignment | None]:
    """Greedy assignment: each registration → earliest (slot, room) without
    a same-student conflict, until rooms in the slot are full. Returns one
    Assignment per input registration (or ``None`` if capacity ran out)."""

    rooms = list(rooms or DEFAULT_ROOMS)
    slots = list(slots or DEFAULT_SLOTS)
    # used[slot_index] = {student_id: True} ; load[slot_index][room] = seats_used
    used: list[dict[str, bool]] = [dict() for _ in slots]
    load: list[dict[str, int]] = [{room: 0 for room in rooms} for _ in slots]

    assignments: list[Assignment | None] = []
    for registration in registrations:
        student = str(registration.get(student_key, "") or "")
        placed: Assignment | None = None
        for slot_index, slot in enumerate(slots):
            if student and student in used[slot_index]:
                continue
            for room in rooms:
                if load[slot_index][room] < seats_per_room:
                    load[slot_index][room] += 1
                    if student:
                        used[slot_index][student] = True
                    placed = Assignment(room, slot.date_label, slot.start, slot.end)
                    break
            if placed:
                break
        assignments.append(placed)
    return assignments


def detect_exam_schedule_request(prompt: str) -> bool:
    text = str(prompt or "")
    return ("教室" in text or "考场" in text or "考点" in text) and (
        "考试" in text or "省考" in text or "排" in text
    )
