const repositoryUrl = 'https://github.com/drghassanabukar/modern-hillbilly-tabs-365';
const releaseApi = 'https://api.github.com/repos/drghassanabukar/modern-hillbilly-tabs-365/releases/latest';
const releasesUrl = repositoryUrl + '/releases/latest';
const cacheKey = 'modern-hillbilly-tabs-365:release-check:v1';
const backoffKey = cacheKey + ':retry';
const cacheLifetime = 60 * 60 * 1000;
const minimumInterval = 60 * 1000;

export interface IUpdateNoticeOptions {
  alwaysShow: boolean;
  enabled: boolean;
  fetcher?: typeof fetch;
  now?: () => number;
  reload?: () => void;
}

interface IRelease {
  tag: string;
  version: string;
  releaseUrl: string;
  downloadUrl: string;
}

interface IReleaseCache {
  checkedAt: number;
  release: IRelease;
}

interface IReleasePayload {
  tag_name?: string;
  html_url?: string;
  draft?: boolean;
  prerelease?: boolean;
  assets?: { name?: string; state?: string; browser_download_url?: string }[];
}

let memoryCache: IReleaseCache | undefined;
let inFlight: Promise<IRelease> | undefined;
let nextRequestAt = 0;

function versionParts(value: string): number[] | undefined {
  if (typeof value !== 'string' || !/^v?\d+\.\d+\.\d+(?:\.\d+)?$/i.test(value) || value.length > 50) return undefined;
  const parts = value.replace(/^v/i, '').split('.').map(Number);
  if (parts.some((part) => !Number.isSafeInteger(part) || part < 0)) return undefined;
  while (parts.length < 4) parts.push(0);
  return parts;
}

function compareVersions(left: string, right: string): number {
  const a = versionParts(left);
  const b = versionParts(right);
  for (let i = 0; i < 4; i++) {
    if (a[i] !== b[i]) return a[i] > b[i] ? 1 : -1;
  }
  return 0;
}

function verifiedUrl(value: string | undefined, expected: string): boolean {
  if (typeof value !== 'string') return false;
  try {
    const url = new URL(value);
    return url.href === expected && url.protocol === 'https:' && !url.username && !url.password;
  } catch { return false; }
}

function parseRelease(payload: IReleasePayload): IRelease {
  if (!payload || payload.draft !== false || payload.prerelease !== false || !versionParts(payload.tag_name)) {
    throw new Error('The latest release does not have a supported stable version. Open releases for details.');
  }
  const releaseUrl = repositoryUrl + '/releases/tag/' + encodeURIComponent(payload.tag_name);
  const downloadUrl = repositoryUrl + '/releases/download/' + encodeURIComponent(payload.tag_name) + '/modern-hillbilly-tabs.sppkg';
  const asset = Array.isArray(payload.assets) && payload.assets.find((item) => item && item.name === 'modern-hillbilly-tabs.sppkg' && item.state === 'uploaded');
  if (!verifiedUrl(payload.html_url, releaseUrl) || !asset || !verifiedUrl(asset.browser_download_url, downloadUrl)) {
    throw new Error('A verified installation package is not available for the latest release. Open releases for details.');
  }
  return { tag: payload.tag_name, version: payload.tag_name.replace(/^v/i, ''), releaseUrl, downloadUrl };
}

function readCache(now: number): IRelease | undefined {
  try {
    const savedBackoff = Number(window.sessionStorage.getItem(backoffKey));
    if (Number.isFinite(savedBackoff) && savedBackoff > now && savedBackoff <= now + 24 * cacheLifetime) nextRequestAt = Math.max(nextRequestAt, savedBackoff);
    if (!memoryCache) {
      const saved = JSON.parse(window.sessionStorage.getItem(cacheKey) || 'null') as IReleaseCache;
      if (saved && typeof saved.checkedAt === 'number' && saved.release) {
        const release = parseRelease({
          tag_name: saved.release.tag, html_url: saved.release.releaseUrl, draft: false, prerelease: false,
          assets: [{ name: 'modern-hillbilly-tabs.sppkg', state: 'uploaded', browser_download_url: saved.release.downloadUrl }]
        });
        memoryCache = { checkedAt: saved.checkedAt, release };
      }
    }
  } catch { /* Private browsing and blocked storage must not prevent update checks. */ }
  return memoryCache && memoryCache.checkedAt <= now && now - memoryCache.checkedAt < cacheLifetime ? memoryCache.release : undefined;
}

function saveBackoff(value: number): void {
  nextRequestAt = value;
  try { window.sessionStorage.setItem(backoffKey, String(value)); } catch { /* Keep the in-memory backoff. */ }
}

function retryTime(response: Response, now: number): number {
  const seconds = Number(response.headers.get('retry-after'));
  const reset = Number(response.headers.get('x-ratelimit-reset')) * 1000;
  return Math.min(now + 24 * cacheLifetime, Math.max(now + minimumInterval, seconds > 0 ? now + seconds * 1000 : 0, reset || 0));
}

async function requestRelease(fetcher: typeof fetch, now: () => number): Promise<IRelease> {
  const controller = new AbortController();
  let timer: number;
  try {
    const timeout = new Promise<IRelease>((_, reject) => {
      timer = window.setTimeout(() => {
        controller.abort();
        reject(new Error('The update check timed out. Try again later or open releases.'));
      }, 5000);
    });
    const request = fetcher(releaseApi, {
      method: 'GET', credentials: 'omit', referrerPolicy: 'no-referrer', cache: 'no-store',
      headers: { Accept: 'application/vnd.github+json' }, signal: controller.signal
    }).then(async (response): Promise<IRelease> => {
      if (response.status === 403 || response.status === 429) {
        saveBackoff(retryTime(response, now()));
        throw new Error('Update checks are temporarily rate limited. Wait before checking again, or open releases.');
      }
      if (!response.ok) throw new Error('Updates could not be checked right now. Try again later or open releases.');
      return parseRelease(await response.json() as IReleasePayload);
    });
    const release = await Promise.race([request, timeout]);
    memoryCache = { checkedAt: now(), release };
    try { window.sessionStorage.setItem(cacheKey, JSON.stringify(memoryCache)); } catch { /* The in-memory cache is sufficient. */ }
    return release;
  } finally { window.clearTimeout(timer); }
}

function latestRelease(force: boolean, options: IUpdateNoticeOptions): Promise<IRelease> {
  const now = options.now || Date.now;
  const cached = readCache(now());
  if (!force && cached) return Promise.resolve(cached);
  if (inFlight) return inFlight;
  if (now() < nextRequestAt) return Promise.reject(new Error('Please wait at least a minute between update checks, or open releases.'));
  const fetcher = options.fetcher || (typeof window.fetch === 'function' ? window.fetch.bind(window) : undefined);
  if (!fetcher) return Promise.reject(new Error('Update checks are unavailable in this browser. Open releases for details.'));
  saveBackoff(now() + minimumInterval);
  inFlight = requestRelease(fetcher, now).then((release) => { inFlight = undefined; return release; }, (error) => { inFlight = undefined; throw error; });
  return inFlight;
}

export class UpdateNotice {
  private disposed = false;
  private checking = false;
  private latest: IRelease | undefined;
  private message = '';
  private state = 'checking';
  private readonly handleKeyDown = (event: KeyboardEvent): void => {
    if (event.key !== 'Enter' && event.key !== ' ' && event.key !== 'Spacebar' && event.key !== 'Space') return;
    const target = event.target as HTMLElement;
    const button = target.closest<HTMLButtonElement>('button.hillbilly-update-button');
    if (!button || !this.host.contains(button)) return;
    event.preventDefault();
    event.stopPropagation();
    if (!this.disposed && !event.repeat && !button.disabled) button.click();
  };
  private readonly handleClick = (event: MouseEvent): void => {
    const target = event.target as HTMLElement;
    const action = target.closest('button.hillbilly-update-button, a.hillbilly-update-button');
    if (!this.disposed && action && this.host.contains(action)) event.stopPropagation();
  };

  public constructor(private readonly host: HTMLElement, private readonly loadedVersion: string, private readonly options: IUpdateNoticeOptions) {
    this.host.classList.add('hillbilly-update-notice');
    this.host.addEventListener('keydown', this.handleKeyDown);
    this.host.addEventListener('click', this.handleClick);
    this.render();
    if (options.enabled) this.check(false);
    else {
      this.state = 'disabled';
      this.message = 'Update checks are turned off.';
      this.render();
    }
  }

  private check(force: boolean): void {
    if (this.disposed || this.checking || !this.options.enabled) return;
    if (!versionParts(this.loadedVersion)) {
      this.state = 'error';
      this.message = 'The loaded version could not be determined. Open releases to check for updates.';
      this.render();
      return;
    }
    this.checking = true;
    this.state = 'checking';
    this.message = 'Checking for updates…';
    this.render();
    latestRelease(force, this.options).then((release) => {
      if (this.disposed) return;
      this.latest = release;
      this.checking = false;
      const newer = compareVersions(release.version, this.loadedVersion) > 0;
      this.state = newer ? 'available' : 'current';
      this.message = newer ? 'Version ' + release.version + ' is available. Upload and enable the new installation package in the SharePoint App Catalog.'
        : 'No newer published version is available. Latest release: ' + release.version + '.';
      this.render();
    }).catch((error: Error) => {
      if (this.disposed) return;
      this.checking = false;
      this.state = 'error';
      this.message = error.message || 'Updates could not be checked. Open releases for details.';
      this.render();
    });
  }

  private link(text: string, url: string): HTMLAnchorElement {
    const link = document.createElement('a');
    link.className = 'hillbilly-update-button';
    link.textContent = text;
    link.href = url;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    return link;
  }

  private render(): void {
    if (this.disposed) return;
    const available = !!this.latest && !!versionParts(this.loadedVersion) && compareVersions(this.latest.version, this.loadedVersion) > 0;
    this.host.hidden = !this.options.alwaysShow && !available;
    this.host.dataset.state = available ? 'available' : this.state;
    this.host.innerHTML = '';
    const version = document.createElement('p');
    version.className = 'hillbilly-update-version';
    version.textContent = 'Loaded version: ' + this.loadedVersion;
    const status = document.createElement('p');
    status.className = 'hillbilly-update-status';
    status.setAttribute('role', 'status');
    status.textContent = this.message;
    const actions = document.createElement('div');
    actions.className = 'hillbilly-update-actions';
    const check = document.createElement('button');
    check.type = 'button';
    check.className = 'hillbilly-update-button';
    check.textContent = this.checking ? 'Checking…' : 'Check for updates';
    check.disabled = this.checking || !this.options.enabled;
    check.addEventListener('click', () => this.check(true));
    actions.appendChild(check);
    if (available) actions.appendChild(this.link('Download update', this.latest.downloadUrl));
    actions.appendChild(this.link(available ? 'Release notes' : 'Open releases', this.latest?.releaseUrl || releasesUrl));
    this.host.appendChild(version);
    this.host.appendChild(status);
    this.host.appendChild(actions);
    if (available) {
      const instruction = document.createElement('p');
      instruction.className = 'hillbilly-update-status';
      instruction.textContent = 'Reloading does not install an update. If an administrator has already deployed the newer package, reload to check for the new assets.'
        + (this.options.alwaysShow ? ' Save or publish your page changes before reloading.' : '');
      const reload = document.createElement('button');
      reload.type = 'button';
      reload.className = 'hillbilly-update-button';
      reload.textContent = 'Already deployed? Reload page';
      reload.addEventListener('click', () => { if (!this.disposed) (this.options.reload || (() => window.location.reload()))(); });
      this.host.appendChild(instruction);
      actions.appendChild(reload);
    }
  }

  public dispose(): void {
    this.disposed = true;
    this.host.removeEventListener('keydown', this.handleKeyDown);
    this.host.removeEventListener('click', this.handleClick);
  }
}
