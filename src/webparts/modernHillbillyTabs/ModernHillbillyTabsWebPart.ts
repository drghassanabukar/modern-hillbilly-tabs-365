import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { CanvasTabs, findControls, ITabData, matchesId } from './CanvasTabs';
import styles from './ModernHillbillyTabsWebPart.module.scss';
import * as strings from 'ModernHillbillyTabsWebPartStrings';
import './Tabs.css';

export interface IModernHillbillyTabsWebPartProps {
  description: string;
  sectionClass: string;
  webpartClass: string;
  tabData: ITabData[];
}

export default class ModernHillbillyTabsWebPart extends BaseClientSideWebPart<IModernHillbillyTabsWebPartProps> {
  private tabs: CanvasTabs | undefined;

  public render(): void {
    this.tabs?.dispose();
    this.tabs = undefined;
    this.domElement.innerHTML = '';
    if (this.displayMode === DisplayMode.Read) {
      this.tabs = new CanvasTabs(this.domElement, this.properties.tabData, this.properties, this.instanceId);
    } else {
      this.domElement.innerHTML = `
        <div class="${styles.modernHillbillyTabs}">
          <div class="${styles.container}">
            <div class="${styles.row}">
              <h2>Modern Hillbilly Tabs 365</h2>
              <p>Place web parts into tabs.</p>
              <ol>
                <li>Place this web part and the content web parts in the same standard page section.</li>
                <li>Open this web part's properties and select <strong>Manage Tab Labels</strong>.</li>
                <li>Select each web part, enter its tab label, and publish the page.</li>
              </ol>
              <p>Content stays visible while you edit the page. The section and web part class settings are optional compatibility overrides.</p>
              <h3>About</h3>
              <p>SharePoint compatibility update by <strong>Dr. Ghassan Abukar, Ph.D.</strong></p>
              <p>Original Modern Hillbilly Tabs project by Mark Rackley.</p>
              <a href="https://github.com/mrackley/Modern_Hillbilly_Tabs" class="${styles.button}">Original project on GitHub</a>
            </div>
          </div>
        </div>`;
    }
  }

  protected onDispose(): void {
    this.tabs?.dispose();
    this.tabs = undefined;
  }

  protected get dataVersion(): Version {
    // Keep the stored property schema compatible with existing pages.
    return Version.parse('1.0');
  }

  private getZones(): { key: string; text: string }[] {
    const controls = findControls(this.domElement, this.properties);
    const options = controls.map((control, index) => {
      // Keep the stored key when a legacy wrapper id maps to a feature-instance id.
      const saved = (this.properties.tabData || []).find((tab) => tab && matchesId(control, tab.WebPartID));
      return { key: saved ? saved.WebPartID : control.id, text: 'Web Part ' + (index + 1) + (control.title ? ' — ' + control.title : '') };
    });
    (this.properties.tabData || []).forEach((tab) => {
      if (tab && tab.WebPartID && !options.some((option) => option.key === tab.WebPartID)) {
        options.push({ key: tab.WebPartID, text: (tab.TabLabel || tab.WebPartID) + ' (unavailable in this section)' });
      }
    });
    return options;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Modern Hillbilly Tabs 365 — compatibility update by Dr. Ghassan Abukar, Ph.D.' },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyFieldCollectionData('tabData', {
              key: 'tabData', label: strings.TabLabels, panelHeader: 'Specify Labels for Tabs',
              manageBtnLabel: 'Manage Tab Labels', value: this.properties.tabData,
              fields: [{
                id: 'WebPartID', title: 'Web Part', type: CustomCollectionFieldType.dropdown,
                required: true, options: this.getZones()
              }, {
                id: 'TabLabel', title: 'Tab Label', type: CustomCollectionFieldType.string
              }],
              disabled: false
            }),
            PropertyPaneTextField('sectionClass', {
              label: strings.SectionClass,
              description: 'Optional section class override. Standard SharePoint markup is detected automatically.'
            }),
            PropertyPaneTextField('webpartClass', {
              label: strings.WebPartClass,
              description: 'Optional web part class override. Standard SharePoint markup is detected automatically.'
            })
          ]
        }]
      }]
    };
  }
}
