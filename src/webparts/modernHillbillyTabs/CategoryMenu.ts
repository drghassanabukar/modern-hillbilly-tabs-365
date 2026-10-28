/** A disclosure-based category chooser; selection stays owned by the tab controller. */
export class CategoryMenu {
  private readonly details: HTMLDetailsElement;
  private readonly summary: HTMLElement;
  private readonly selected: HTMLElement;
  private readonly count: HTMLElement;
  private readonly search: HTMLInputElement;
  private readonly results: HTMLElement;
  private readonly empty: HTMLElement;

  public constructor(host: HTMLElement, private readonly tablist: HTMLElement, instanceId: string) {
    const wrapper = document.createElement('div');
    wrapper.className = 'hillbilly-category-menu';
    const label = document.createElement('div');
    label.className = 'hillbilly-category-label';
    label.textContent = 'Category';
    this.details = document.createElement('details');
    this.details.className = 'hillbilly-category-details';
    this.summary = document.createElement('summary');
    this.summary.className = 'hillbilly-category-summary';
    this.selected = document.createElement('span');
    this.selected.className = 'hillbilly-category-selected';
    this.count = document.createElement('span');
    this.count.className = 'hillbilly-category-count';
    const chevron = document.createElement('span');
    chevron.className = 'hillbilly-category-chevron';
    chevron.setAttribute('aria-hidden', 'true');
    chevron.textContent = '⌄';
    this.summary.appendChild(this.selected);
    this.summary.appendChild(this.count);
    this.summary.appendChild(chevron);
    const panel = document.createElement('div');
    panel.className = 'hillbilly-category-panel';
    const searchLabel = document.createElement('label');
    searchLabel.className = 'hillbilly-category-search-label';
    searchLabel.htmlFor = instanceId + '-category-search';
    searchLabel.textContent = 'Find a category';
    this.search = document.createElement('input');
    this.search.type = 'search';
    this.search.id = searchLabel.htmlFor;
    this.search.className = 'hillbilly-category-search';
    this.search.placeholder = 'Search categories…';
    this.search.autocomplete = 'off';
    this.results = document.createElement('p');
    this.results.className = 'hillbilly-category-results';
    this.results.setAttribute('role', 'status');
    this.empty = document.createElement('p');
    this.empty.className = 'hillbilly-category-empty';
    this.empty.textContent = 'Try a different word or clear the search.';
    tablist.setAttribute('aria-orientation', 'vertical');
    panel.appendChild(searchLabel);
    panel.appendChild(this.search);
    panel.appendChild(this.results);
    panel.appendChild(tablist);
    panel.appendChild(this.empty);
    this.details.appendChild(this.summary);
    this.details.appendChild(panel);
    wrapper.appendChild(label);
    wrapper.appendChild(this.details);
    host.appendChild(wrapper);

    this.search.addEventListener('input', () => this.filter());
    this.search.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp' && event.key !== 'Enter') return;
      // Enter must not submit a surrounding page form, including an empty search.
      event.preventDefault();
      const available = this.buttons().filter((button) => !button.hidden && !button.disabled);
      if (!available.length) return;
      if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
        available[event.key === 'ArrowDown' ? 0 : available.length - 1].focus();
      } else if (event.key === 'Enter') {
        available[0].click();
      }
    });
    tablist.addEventListener('click', (event) => {
      const button = (event.target as HTMLElement).closest<HTMLButtonElement>('[role="tab"]');
      if (button && !button.disabled) this.close();
    });
    this.details.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        this.close();
      }
    });
    this.details.addEventListener('toggle', () => {
      if (this.details.open) this.search.focus();
      else if (this.search.value) {
        this.search.value = '';
        this.filter();
      }
    });
  }

  private buttons(): HTMLButtonElement[] {
    return Array.from(this.tablist.querySelectorAll<HTMLButtonElement>('[role="tab"]'));
  }

  private filter(): void {
    const query = this.search.value.trim().toLocaleLowerCase();
    const buttons = this.buttons();
    const visible = buttons.filter((button) => {
      const match = (button.textContent || '').toLocaleLowerCase().indexOf(query) >= 0;
      if (button.hidden !== !match) button.hidden = !match;
      return match;
    });
    const available = visible.filter((button) => !button.disabled);
    const focusable = available.find((button) => button.getAttribute('aria-selected') === 'true') || available[0];
    buttons.forEach((button) => { button.tabIndex = button === focusable ? 0 : -1; });
    if (this.empty.hidden !== !!visible.length) this.empty.hidden = !!visible.length;
    const resultText = !query ? '' : !visible.length ? 'No categories match your search.'
      : visible.length + (visible.length === 1 ? ' category found' : ' categories found');
    if (this.results.textContent !== resultText) this.results.textContent = resultText;
    if (this.results.hidden !== !query) this.results.hidden = !query;
  }

  public update(selectedLabel: string): void {
    const label = selectedLabel || 'Choose a category';
    if (this.selected.textContent !== label) this.selected.textContent = label;
    this.selected.title = label;
    const count = this.buttons().length;
    const countText = count + (count === 1 ? ' category' : ' categories');
    if (this.count.textContent !== countText) this.count.textContent = countText;
    this.filter();
  }

  private close(): void {
    this.details.open = false;
    this.search.value = '';
    this.filter();
    this.summary.focus();
  }
}
