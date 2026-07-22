// Type-ahead enhancement for <select> elements. Wraps a native select in a
// text input + filtered dropdown: typing filters the options by substring,
// arrow keys navigate, Enter/click selects. The underlying select stays in
// the DOM (hidden) and remains the source of truth — we write sel.value and
// dispatch a "change" event, so existing listeners and code that reads the
// select by id keep working unmodified. Option changes (re-population) and
// the disabled attribute are tracked with a MutationObserver.

/**
 * Returns the wrapper element containing the input, dropdown, and the (now
 * hidden) select. For selects already in the DOM the wrapper is swapped in
 * place; for detached selects (built during grid rendering) append the
 * RETURNED wrapper — appending the select itself would tear it back out.
 */
export function enhanceSelect(sel: HTMLSelectElement): HTMLElement {
  if (sel.dataset.comboboxed === "1") return sel.parentElement as HTMLElement;
  sel.dataset.comboboxed = "1";

  const wrap = document.createElement("span");
  wrap.className = "combobox";
  wrap.style.cssText = "position:relative;display:inline-block;width:100%;";

  const input = document.createElement("input");
  input.type = "text";
  input.autocomplete = "off";
  input.spellcheck = false;
  input.setAttribute("role", "combobox");
  input.setAttribute("aria-expanded", "false");

  const list = document.createElement("div");
  list.className = "combobox-list";
  list.style.cssText =
    "position:absolute;left:0;right:0;top:100%;z-index:1000;display:none;" +
    "max-height:180px;overflow-y:auto;background:#fff;border:1px solid #8a8886;" +
    "border-top:none;box-shadow:0 4px 8px rgba(0,0,0,.15);font-size:12px;";

  sel.parentNode?.insertBefore(wrap, sel);
  wrap.append(input, list);
  wrap.appendChild(sel);
  sel.style.display = "none";

  let highlighted = -1;
  let filtered: HTMLOptionElement[] = [];

  const options = (): HTMLOptionElement[] => [...sel.options];

  const currentLabel = (): string => {
    const o = options().find((o) => o.value === sel.value);
    return o?.text ?? "";
  };

  const syncFromSelect = (): void => {
    input.value = currentLabel();
    input.placeholder = options().find((o) => !o.value)?.text ?? "Type to filter…";
    input.disabled = sel.disabled;
  };

  const close = (): void => {
    list.style.display = "none";
    input.setAttribute("aria-expanded", "false");
    highlighted = -1;
  };

  const pick = (o: HTMLOptionElement): void => {
    sel.value = o.value;
    close();
    syncFromSelect();
    sel.dispatchEvent(new Event("change", { bubbles: true }));
  };

  const render = (term: string): void => {
    const t = term.trim().toLowerCase();
    filtered = options().filter(
      (o) => o.value && (!t || o.text.toLowerCase().includes(t) || o.value.toLowerCase().includes(t))
    );
    list.innerHTML = "";
    if (filtered.length === 0) {
      const empty = document.createElement("div");
      empty.textContent = "No matches";
      empty.style.cssText = "padding:4px 8px;color:#605e5c;";
      list.appendChild(empty);
    }
    filtered.forEach((o, i) => {
      const item = document.createElement("div");
      item.textContent = o.text;
      item.style.cssText = "padding:4px 8px;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;";
      if (i === highlighted) item.style.background = "#edebe9";
      if (o.value === sel.value) item.style.fontWeight = "600";
      // mousedown, not click: fires before the input's blur closes the list
      item.addEventListener("mousedown", (e) => {
        e.preventDefault();
        pick(o);
      });
      item.addEventListener("mousemove", () => {
        if (highlighted !== i) {
          highlighted = i;
          render(input.value === currentLabel() ? "" : input.value);
        }
      });
      list.appendChild(item);
    });
    list.style.display = "";
    input.setAttribute("aria-expanded", "true");
  };

  input.addEventListener("focus", () => {
    input.select();
    highlighted = -1;
    render(""); // show all on focus; filter kicks in once they type
  });

  input.addEventListener("input", () => {
    highlighted = 0;
    render(input.value);
  });

  input.addEventListener("keydown", (e) => {
    if (list.style.display === "none" && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
      render(input.value === currentLabel() ? "" : input.value);
      e.preventDefault();
      return;
    }
    switch (e.key) {
      case "ArrowDown":
        highlighted = Math.min(highlighted + 1, filtered.length - 1);
        render(input.value === currentLabel() ? "" : input.value);
        list.children[highlighted]?.scrollIntoView({ block: "nearest" });
        e.preventDefault();
        break;
      case "ArrowUp":
        highlighted = Math.max(highlighted - 1, 0);
        render(input.value === currentLabel() ? "" : input.value);
        list.children[highlighted]?.scrollIntoView({ block: "nearest" });
        e.preventDefault();
        break;
      case "Enter":
        if (highlighted >= 0 && filtered[highlighted]) pick(filtered[highlighted]);
        else if (filtered.length === 1) pick(filtered[0]);
        e.preventDefault();
        break;
      case "Escape":
        close();
        syncFromSelect();
        break;
    }
  });

  input.addEventListener("blur", () => {
    close();
    syncFromSelect(); // revert partial typing to the actual selection
  });

  // Re-populated options or toggled disabled → refresh the input. Deferred a
  // tick so code that repopulates options and then sets sel.value in the same
  // synchronous block is observed in its final state.
  new MutationObserver(() => setTimeout(syncFromSelect, 0)).observe(sel, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ["disabled"],
  });
  // Programmatic change events (and our own picks) also refresh the display.
  sel.addEventListener("change", syncFromSelect);

  syncFromSelect();
  return wrap;
}
