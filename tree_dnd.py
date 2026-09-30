"""
Trascinamento con il mouse (drag & drop) per riordinare le righe di un ttk.Treeview.

Uso:
    TreeDragDrop(tree, theme,
                 can_drag=lambda iid: True,
                 can_drop=lambda src, target, where: True,
                 on_drop=lambda src, target, where: ...)

`where` vale "before", "after" oppure "inside" (solo se allow_inside(target) e' vero:
ad esempio per trascinare uno step dentro una ripetizione o un allenamento dentro una settimana).
"""
import time
import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional, Tuple

DRAG_THRESHOLD = 6  # pixel prima di considerare il click un trascinamento


class TreeDragDrop:
    def __init__(self, tree: ttk.Treeview, theme,
                 on_drop: Callable[[str, str, str], None],
                 can_drag: Callable[[str], bool] = lambda iid: True,
                 can_drop: Callable[[str, str, str], bool] = lambda s, t, w: True,
                 allow_inside: Callable[[str], bool] = lambda iid: False):
        self.tree = tree
        self.theme = theme
        self.on_drop = on_drop
        self.can_drag = can_drag
        self.can_drop = can_drop
        self.allow_inside = allow_inside
        self._src: Optional[str] = None
        self._start: Tuple[int, int] = (0, 0)
        self._dragging = False
        self._target: Optional[Tuple[str, str]] = None
        self._line = tk.Frame(tree, height=3)
        self._box = [tk.Frame(tree) for _ in range(4)]  # cornice per "inside"
        self._ghost: Optional[tk.Toplevel] = None
        self._last_drop = 0.0
        tree.bind("<ButtonPress-1>", self._press, add="+")
        tree.bind("<B1-Motion>", self._motion)
        tree.bind("<ButtonRelease-1>", self._release, add="+")
        tree.bind("<Escape>", lambda e: self._cancel(), add="+")

    # ------------------------------------------------------------------
    def _press(self, e):
        self._src = None
        self._dragging = False
        if self.tree.identify_region(e.x, e.y) not in ("tree", "cell"):
            return
        iid = self.tree.identify_row(e.y)
        if iid and self.can_drag(iid):
            self._src = iid
            self._start = (e.x, e.y)

    def _motion(self, e):
        if self._src is None:
            return None  # comportamento normale del Treeview (selezione)
        if not self._dragging:
            if abs(e.y - self._start[1]) < DRAG_THRESHOLD and abs(e.x - self._start[0]) < DRAG_THRESHOLD:
                return "break"
            self._dragging = True
            self.tree.configure(cursor="fleur")
            self._show_ghost()
        self._move_ghost(e)
        self._autoscroll(e.y)
        self._target = self._compute_target(e.y)
        self._draw_indicator()
        return "break"  # niente selezione a intervallo mentre si trascina

    def _release(self, e):
        if self._dragging:
            self._last_drop = time.monotonic()
            target = self._target
            src = self._src
            self._cancel()
            if target and src:
                self.on_drop(src, target[0], target[1])
            return "break"
        self._src = None
        return None

    def recently_dropped(self, seconds: float = 0.6) -> bool:
        """True subito dopo un rilascio: evita che due click ravvicinati diventino un doppio clic."""
        return time.monotonic() - self._last_drop < seconds

    def _cancel(self):
        self._dragging = False
        self._src = None
        self._target = None
        self.tree.configure(cursor="")
        self._line.place_forget()
        for f in self._box:
            f.place_forget()
        if self._ghost is not None:
            self._ghost.destroy()
            self._ghost = None

    # ------------------------------------------------------------------
    def _compute_target(self, y) -> Optional[Tuple[str, str]]:
        iid = self.tree.identify_row(y)
        if not iid:
            # sotto l'ultima riga: dopo l'ultimo elemento visibile di primo livello
            kids = self.tree.get_children("")
            if not kids:
                return None
            last = kids[-1]
            while self.tree.item(last, "open") and self.tree.get_children(last):
                last = self.tree.get_children(last)[-1]
            iid, where = last, "after"
        else:
            bbox = self.tree.bbox(iid)
            if not bbox:
                return None
            _, top, _, h = bbox
            rel = (y - top) / max(h, 1)
            if self.allow_inside(iid) and 0.25 < rel < 0.75:
                where = "inside"
            else:
                where = "before" if rel < 0.5 else "after"
        if iid == self._src or not self.can_drop(self._src, iid, where):
            return None
        return iid, where

    def _draw_indicator(self):
        c = self.theme.c
        self._line.place_forget()
        for f in self._box:
            f.place_forget()
        if not self._target:
            return
        iid, where = self._target
        bbox = self.tree.bbox(iid)
        if not bbox:
            return
        x, y, w, h = bbox
        width = self.tree.winfo_width() - 4
        if where == "inside":
            for f in self._box:
                f.configure(bg=c["accent"])
            self._box[0].place(x=2, y=y, width=width, height=2)
            self._box[1].place(x=2, y=y + h - 2, width=width, height=2)
            self._box[2].place(x=2, y=y, width=2, height=h)
            self._box[3].place(x=width, y=y, width=2, height=h)
        else:
            self._line.configure(bg=c["accent"])
            yy = y - 1 if where == "before" else y + h - 2
            indent = 20 * self._depth(iid)
            self._line.place(x=2 + indent, y=yy, width=width - indent, height=3)

    def _depth(self, iid) -> int:
        d = 0
        p = self.tree.parent(iid)
        while p:
            d += 1
            p = self.tree.parent(p)
        return d

    def _autoscroll(self, y):
        h = self.tree.winfo_height()
        if y < 20:
            self.tree.yview_scroll(-1, "units")
        elif y > h - 20:
            self.tree.yview_scroll(1, "units")

    def _show_ghost(self):
        c = self.theme.c
        text = self.tree.item(self._src, "text").strip()
        vals = self.tree.item(self._src, "values")
        if vals and str(vals[0]).strip():
            text = f"{text}  {vals[0]}"
        self._ghost = tk.Toplevel(self.tree)
        self._ghost.wm_overrideredirect(True)
        try:
            self._ghost.attributes("-alpha", 0.85)
        except tk.TclError:
            pass
        tk.Label(self._ghost, text=text[:60], bg=c["card"], fg=c["fg"], padx=10, pady=4,
                 highlightthickness=1, highlightbackground=c["accent"],
                 font=self.theme.font).pack()

    def _move_ghost(self, e):
        if self._ghost is not None:
            self._ghost.wm_geometry(f"+{e.x_root + 14}+{e.y_root + 10}")
