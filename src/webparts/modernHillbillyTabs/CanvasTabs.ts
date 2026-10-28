/** SharePoint canvas markup is not a public API. Keep compatibility handling here. */
export interface ITabData {
  WebPartID: string;
  TabLabel: string;
}

export interface ICanvasOptions {
  sectionClass?: string;
  webpartClass?: string;
}

export interface ICanvasControl {
  element: HTMLElement;
  id: string;
  aliases: string[];
  title: string;
}

const controlSelector = '.ControlZone, [data-automation-id="CanvasControl"], [data-sp-feature-instance-id]';
const sectionSelector = '.CanvasSection, [data-automation-id="CanvasSection"]';
// Module-local ownership prevents two tab web parts from changing the same control.
const owners = new WeakMap<HTMLElement, CanvasTabs>();

function classSelector(value: string | undefined): string {
  // These settings historically contain a class name, never arbitrary CSS.
  return value && /^[a-zA-Z_][\w-]*$/.test(value) ? '.' + value : '';
}

export function findSection(host: HTMLElement, options: ICanvasOptions): HTMLElement | undefined {
  const custom = classSelector(options.sectionClass);
  return (custom && host.closest<HTMLElement>(custom)) || host.closest<HTMLElement>(sectionSelector) || undefined;
}

function normalizeId(value: string): string {
  const trimmed = String(value || '').trim();
  // Retain exact matching for non-GUID ids; migrate known legacy wrapper prefixes.
  const guid = /^(?:control_|controlzone_|webpart_)?\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?$/i.exec(trimmed);
  return guid ? guid[1].toLowerCase() : trimmed;
}

export function matchesId(control: ICanvasControl, savedId: string): boolean {
  return control.aliases.some((alias) => normalizeId(alias) === normalizeId(savedId));
}

export function findControls(host: HTMLElement, options: ICanvasOptions): ICanvasControl[] {
  const section = findSection(host, options);
  if (!section) return []; // Never fall back to searching the whole page.
  const custom = classSelector(options.webpartClass);
  const selector = controlSelector + (custom ? ', ' + custom : '');
  const elements = Array.from(section.querySelectorAll<HTMLElement>(selector));
  return elements.filter((element) => {
    if (element.contains(host) || host.contains(element)) return false;
    if (findSection(element, options) !== section) return false;
    // Prefer the outer canvas control over the feature wrapper inside it.
    return !elements.some((parent) => parent !== element && parent.contains(element));
  }).map((element) => {
    const identities = [element].concat(Array.from(element.querySelectorAll<HTMLElement>(selector)));
    const aliases: string[] = [];
    let stableId = '';
    identities.forEach((identity) => {
      const featureId = identity.getAttribute('data-sp-feature-instance-id');
      if (featureId) {
        if (!stableId) stableId = featureId;
        aliases.push(featureId);
      }
      if (identity.id) aliases.push(identity.id);
    });
    const heading = element.querySelector<HTMLElement>('[data-sp-web-part-title], h2, h3, [role="heading"]');
    return {
      element,
      id: stableId || element.id,
      aliases,
      title: element.getAttribute('data-sp-web-part-title') || (heading && heading.textContent || '').trim()
    };
  }).filter((control) => !!control.id);
}

interface IOriginalState {
  attributes: { [name: string]: string | null };
  display: string;
  displayPriority: string;
}

export class CanvasTabs {
  private readonly originals = new Map<HTMLElement, IOriginalState>();
  private readonly hiddenDisplays = new Map<HTMLElement, { display: string; priority: string }>();
  private readonly buttons: HTMLButtonElement[] = [];
  private readonly status: HTMLParagraphElement;
  private observer: MutationObserver;
  private observedRoot: HTMLElement;
  private timer: number | undefined;
  private active = 0;
  private disposed = false;
  private readonly tabs: ITabData[];

  public constructor(
    private readonly host: HTMLElement,
    tabData: ITabData[],
    private readonly options: ICanvasOptions,
    private readonly instanceId: string
  ) {
    const seen: string[] = [];
    this.tabs = (Array.isArray(tabData) ? tabData : []).filter((tab) => {
      if (!tab || typeof tab.WebPartID !== 'string' || !tab.WebPartID.trim()) return false;
      const id = normalizeId(tab.WebPartID);
      if (seen.indexOf(id) !== -1) return false;
      seen.push(id);
      return true;
    });
    const tablist = document.createElement('div');
    tablist.className = 'hillbilly-tabs-list';
    tablist.setAttribute('role', 'tablist');
    tablist.setAttribute('aria-label', 'Page content');
    this.tabs.forEach((tab, index) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.id = this.instanceId + '-tab-' + index;
      button.setAttribute('role', 'tab');
      button.textContent = tab.TabLabel || 'Tab ' + (index + 1);
      button.addEventListener('click', () => { this.active = index; this.sync(); });
      button.addEventListener('keydown', (event) => this.onKeyDown(event, index));
      this.buttons.push(button);
      tablist.appendChild(button);
    });
    this.status = document.createElement('p');
    this.status.className = 'hillbilly-tabs-status';
    this.status.setAttribute('role', 'status');
    this.host.appendChild(tablist);
    this.host.appendChild(this.status);
    this.observer = new MutationObserver((records) => {
      // Changes in our own tab/status UI must not trigger an observer loop.
      if (records.every((record) => this.host.contains(record.target))) return;
      if (this.timer === undefined) {
        this.timer = window.setTimeout(() => { this.timer = undefined; this.sync(); }, 0);
      }
    });
    this.sync();
  }

  private sync(): void {
    if (this.disposed) return;
    const section = findSection(this.host, this.options);
    const root = section || document.body;
    if (root !== this.observedRoot) {
      this.observer.disconnect();
      this.observer.observe(root, {
        childList: true, subtree: true, attributes: true,
        attributeFilter: ['id', 'class', 'style', 'hidden', 'data-automation-id', 'data-sp-feature-instance-id']
      });
      this.observedRoot = root;
    }
    const controls = findControls(this.host, this.options);
    const targets = this.tabs.map((tab) => controls.find((control) => matchesId(control, tab.WebPartID)));
    const claimed = new Set<HTMLElement>();
    const usable = targets.map((control) => {
      if (!control || claimed.has(control.element) || (owners.has(control.element) && owners.get(control.element) !== this)) return false;
      claimed.add(control.element);
      return true;
    });
    if (!usable[this.active]) this.active = usable.indexOf(true);
    // Release replaced/removed controls; never move SharePoint-owned elements.
    this.originals.forEach((_, element) => {
      if (!targets.some((control, index) => usable[index] && control.element === element)) this.restore(element);
    });
    targets.forEach((control, index) => {
      const button = this.buttons[index];
      button.disabled = !usable[index];
      button.setAttribute('aria-selected', String(index === this.active));
      button.tabIndex = index === this.active ? 0 : -1;
      if (!usable[index]) { button.removeAttribute('aria-controls'); return; }
      const element = control.element;
      if (!this.originals.has(element)) {
        const attributes: { [name: string]: string | null } = {};
        ['id', 'role', 'aria-labelledby', 'aria-hidden', 'hidden', 'tabindex'].forEach((name) => {
          attributes[name] = element.getAttribute(name);
        });
        this.originals.set(element, { attributes, display: element.style.getPropertyValue('display'), displayPriority: element.style.getPropertyPriority('display') });
        owners.set(element, this);
      }
      if (!element.id) element.id = this.instanceId + '-panel-' + index;
      button.setAttribute('aria-controls', element.id);
      element.setAttribute('role', 'tabpanel');
      element.setAttribute('aria-labelledby', button.id);
      element.setAttribute('tabindex', '0');
      const selected = index === this.active;
      if (element.hidden !== !selected) element.hidden = !selected;
      element.setAttribute('aria-hidden', String(!selected));
      if (selected) {
        const previous = this.hiddenDisplays.get(element);
        // Restore only our own override; leave later active-layout changes alone.
        if (previous) {
          if (previous.display) element.style.setProperty('display', previous.display, previous.priority);
          else element.style.removeProperty('display');
          this.hiddenDisplays.delete(element);
        }
      } else {
        if (!this.hiddenDisplays.has(element)) this.hiddenDisplays.set(element, {
          display: element.style.getPropertyValue('display'), priority: element.style.getPropertyPriority('display')
        });
        if (element.style.getPropertyValue('display') !== 'none' || element.style.getPropertyPriority('display') !== 'important') {
          element.style.setProperty('display', 'none', 'important');
        }
      }
    });
    const missing = usable.filter((value) => !value).length;
    const message = !section ? 'Cannot find this page section. Place tabs and content in the same standard section.'
      : !this.tabs.length ? 'Edit this web part and select Manage Tab Labels to configure tabs.'
      : missing ? 'Some configured web parts are unavailable. If they do not load, check Manage Tab Labels and keep them in this section.' : '';
    if (this.status.textContent !== message) this.status.textContent = message;
    if (this.status.hidden !== !message) this.status.hidden = !message;
  }

  private onKeyDown(event: KeyboardEvent, index: number): void {
    const available = this.buttons.map((button, i) => button.disabled ? -1 : i).filter((i) => i >= 0);
    let next = available.indexOf(index);
    if (event.key === 'ArrowRight') next = (next + 1) % available.length;
    else if (event.key === 'ArrowLeft') next = (next + available.length - 1) % available.length;
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = available.length - 1;
    else return;
    event.preventDefault();
    this.active = available[next];
    this.sync();
    this.buttons[this.active].focus();
  }

  private restore(element: HTMLElement): void {
    const original = this.originals.get(element);
    Object.keys(original.attributes).forEach((name) => {
      const value = original.attributes[name];
      if (value === null) element.removeAttribute(name);
      else element.setAttribute(name, value);
    });
    if (original.display) element.style.setProperty('display', original.display, original.displayPriority);
    else element.style.removeProperty('display');
    owners.delete(element);
    this.hiddenDisplays.delete(element);
    this.originals.delete(element);
  }

  public dispose(): void {
    this.disposed = true;
    this.observer.disconnect();
    if (this.timer !== undefined) window.clearTimeout(this.timer);
    this.originals.forEach((_, element) => this.restore(element));
  }
}
