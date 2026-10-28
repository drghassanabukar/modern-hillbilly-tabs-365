import { CategoryMenu } from './CategoryMenu';

export type TabDesign = 'underline' | 'pills' | 'outlined' | 'segmented' | 'blocks' | 'minimal';
export type TabAlignment = 'start' | 'center' | 'stretch';
export type TabLayout = 'tabs' | 'menu' | 'auto';

export interface ITabAppearanceOptions {
  tabDesign?: string;
  tabAlignment?: string;
  tabLayout?: string;
}

export const tabDesignOptions: { key: TabDesign; text: string }[] = [
  { key: 'underline', text: 'Classic underline' },
  { key: 'pills', text: 'Rounded pills' },
  { key: 'outlined', text: 'Outlined tabs' },
  { key: 'segmented', text: 'Segmented control' },
  { key: 'blocks', text: 'Solid blocks' },
  { key: 'minimal', text: 'Minimal' }
];

export const tabAlignmentOptions: { key: TabAlignment; text: string }[] = [
  { key: 'start', text: 'Left' },
  { key: 'center', text: 'Centered' },
  { key: 'stretch', text: 'Full width' }
];

export const tabLayoutOptions: { key: TabLayout; text: string }[] = [
  { key: 'tabs', text: 'Tab bar' },
  { key: 'menu', text: 'Searchable category menu' },
  { key: 'auto', text: 'Automatic (menu for more than 8 categories)' }
];

export function getTabAppearance(options: ITabAppearanceOptions): { design: TabDesign; alignment: TabAlignment } {
  return {
    design: tabDesignOptions.find((option) => option.key === options.tabDesign)?.key || 'underline',
    alignment: tabAlignmentOptions.find((option) => option.key === options.tabAlignment)?.key || 'start'
  };
}

export function createTabList(options: ITabAppearanceOptions, label: string): HTMLDivElement {
  const appearance = getTabAppearance(options);
  const tablist = document.createElement('div');
  tablist.className = 'hillbilly-tabs-list';
  tablist.dataset.tabDesign = appearance.design;
  tablist.dataset.tabAlignment = appearance.alignment;
  tablist.setAttribute('role', 'tablist');
  tablist.setAttribute('aria-label', label);
  return tablist;
}

export function getTabLayout(options: ITabAppearanceOptions, count: number): 'tabs' | 'menu' {
  return options.tabLayout === 'menu' || (options.tabLayout === 'auto' && count > 8) ? 'menu' : 'tabs';
}

export function appendTabNavigation(host: HTMLElement, tablist: HTMLElement, options: ITabAppearanceOptions, instanceId: string, count: number): CategoryMenu | undefined {
  if (getTabLayout(options, count) === 'menu') return new CategoryMenu(host, tablist, instanceId);
  host.appendChild(tablist);
  return undefined;
}

/** Preview buttons never discover or change SharePoint canvas controls. */
export function renderTabPreview(host: HTMLElement, options: ITabAppearanceOptions, instanceId: string, categoryCount: number = 3): void {
  const tablist = createTabList(options, 'Tab design preview');
  const panel = document.createElement('div');
  panel.id = instanceId + '-preview-panel';
  panel.className = 'hillbilly-tabs-preview-panel';
  panel.setAttribute('role', 'tabpanel');
  panel.tabIndex = 0;
  const labels = ['Overview', 'Documents', 'Resources'];
  const buttons: HTMLButtonElement[] = [];
  const select = (index: number): void => {
    buttons.forEach((button, i) => {
      button.setAttribute('aria-selected', String(i === index));
      button.tabIndex = i === index ? 0 : -1;
    });
    panel.setAttribute('aria-labelledby', buttons[index].id);
    panel.textContent = labels[index] + ' content preview';
    menu?.update(labels[index]);
  };
  labels.forEach((label, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.id = instanceId + '-preview-tab-' + index;
    button.textContent = label;
    button.setAttribute('role', 'tab');
    button.setAttribute('aria-controls', panel.id);
    button.addEventListener('click', () => select(index));
    button.addEventListener('keydown', (event) => {
      let next: number;
      const visible = buttons.map((item, i) => item.hidden ? -1 : i).filter((i) => i >= 0);
      const position = visible.indexOf(index);
      if (event.key === 'ArrowRight' || event.key === 'ArrowDown') next = visible[(position + 1) % visible.length];
      else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') next = visible[(position + visible.length - 1) % visible.length];
      else if (event.key === 'Home') next = visible[0];
      else if (event.key === 'End') next = visible[visible.length - 1];
      else return;
      event.preventDefault();
      select(next);
      buttons[next].focus();
    });
    buttons.push(button);
    tablist.appendChild(button);
  });
  const menu = appendTabNavigation(host, tablist, options, instanceId + '-preview', Math.max(categoryCount, labels.length));
  host.appendChild(panel);
  select(0);
}

export function getAccentTextColor(accent: string): string {
  let hex = accent.replace('#', '');
  if (/^[0-9a-f]{3}$/i.test(hex)) hex = hex.split('').map((value) => value + value).join('');
  if (!/^[0-9a-f]{6}$/i.test(hex)) return '#ffffff';
  const channels = [0, 2, 4].map((offset) => {
    const value = parseInt(hex.substr(offset, 2), 16) / 255;
    return value <= 0.04045 ? value / 12.92 : Math.pow((value + 0.055) / 1.055, 2.4);
  });
  const luminance = channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722;
  return luminance > 0.179 ? '#000000' : '#ffffff';
}
