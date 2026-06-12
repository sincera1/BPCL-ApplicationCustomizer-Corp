import { Log } from '@microsoft/sp-core-library';
import styles from './BpclApplicationCustomizer.module.scss';
import MenuService, { IMenuItem } from './Service';
import { SPPermission } from '@microsoft/sp-page-context';

import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName} from '@microsoft/sp-application-base';

interface IMenuNode {
  Id: number;
  Title: string;
  SiteURL?: { Url: string };
  children: IMenuNode[];
}

export default class ApplicationCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {

  private _top: PlaceholderContent | undefined;




  public async onInit(): Promise<void> {
    Log.info('ApplicationCustomizer', 'Initialized');

    this._hideAppBar();
    this._loadBootstrapIcons();

    await this._renderTop();

    this._renderBottom();

    this.context.application.navigatedEvent.add(this, () => {
      this._renderBottom();
    });

    return Promise.resolve();
  }


  /* ================= ICONS ================= */
  private _loadBootstrapIcons(): void {
    if (document.getElementById('bootstrap-icons')) return;

    const link = document.createElement('link');
    link.id = 'bootstrap-icons';
    link.rel = 'stylesheet';
    link.href =
      'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css';
    document.head.appendChild(link);
  }

  /* ================= HIDE SHAREPOINT CHROME ================= */
  private _hideAppBar(): void {
    const style = document.createElement('style');
    style.innerHTML = `
      #spSiteHeader { display: none !important; }
      #sp-appBar { display: none !important; }
      .ms-HorizontalNav { margin-left: 0 !important; }
      #spCommandBar { display: none !important; }
      #vpc_Page\\.SiteFooter\\.internal\\.03025612-a400-4804-a78e-e1493200a43b { display: none !important; }
      #CommentsWrapper { display: none !important; }
      #O365_MainLink_Settings {
      display: none !important;
    }

       #spLeftNav {
      display: none !important;
    }
    
    `;
    document.head.appendChild(style);
  }

  /* ================= BUILD LEVEL 0 → LEVEL 1 TREE ================= */
  private _buildMenuTree(items: IMenuItem[], category: string): IMenuNode[] {

    const parents = items.filter(
      i => i.Level === 0 && i.Category === category
    );

    return parents.map(parent => ({
      Id: parent.Id,
      Title: parent.Title,
      SiteURL: parent.SiteURL,
      children: items
        .filter(child =>
          child.Level === 1 &&
          child.ParentIDId === parent.Id
        )
        .map(child => ({
          Id: child.Id,
          Title: child.Title,
          SiteURL: child.SiteURL,
          children: []
        }))
    }));
  }

  /* ================= RENDER DROPDOWN HTML ================= */


  private _renderMenuHtml(menuTree: IMenuNode[]): string {

    const renderItems = (items: IMenuNode[]) => items.map(parent => `
    <li class="${styles.submenuItem}">

      ${parent.children.length > 0
        ? `
          <a class="${styles.submenuLink} d-flex align-items-center justify-content-between">
            ${parent.Title}
            <i class="bi bi-caret-right-fill"></i>
          </a>
        `
        : `
          <a class="${styles.submenuLink}" href="${parent.SiteURL?.Url || '#'}" target="_blank" data-interception="off">
            ${parent.Title}
          </a>
        `
      }

      ${parent.children.length > 0
        ? `
          <ul class="${styles.rightSubmenu}">
            ${parent.children.map(child => `
              <li>
                <a href="${child.SiteURL?.Url || '#'}" target="_blank" data-interception="off" tabindex="0">
                  ${child.Title}
                </a>
              </li>
            `).join('')}
          </ul>
        `
        : ''
      }

    </li>
  `).join('');

    // ✅ If 10 or less → normal single column
    if (menuTree.length <= 10) {
      return renderItems(menuTree);
    }

    // ✅ If more than 10 → split into 2 columns (10 + remaining)
    const firstColumn = menuTree.slice(0, 10);
    const secondColumn = menuTree.slice(10);

    return `
    <li class="${styles.submenuItem} ${styles.submenuColumnsWrapper}">
      
      <div class="${styles.submenuWrapper}">
        
        <ul class="${styles.column}">
          ${renderItems(firstColumn)}
        </ul>

        <ul class="${styles.column}">
          ${renderItems(secondColumn)}
        </ul>

      </div>

    </li>
  `;
  }


  /* ================= TOP NAV ================= */
  private async _renderTop(): Promise<void> {

    const webUrl = this.context.pageContext.web.absoluteUrl;

    const settingsUrl = `${webUrl}/_layouts/15/settings.aspx`;
    const siteContentsUrl = `${webUrl}/_layouts/15/viewlsts.aspx`;

    if (this._top) return;

    this._top = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top
    );
    if (!this._top) return;

    //  PERMISSION CHECK – CORRECT PLACE
    const hasAdminAccess =
      this.context.pageContext.web.permissions.hasPermission(
        SPPermission.manageWeb
      );

    /* ===== LIST / LIBRARY DETECTION ===== */
    const listContext = this.context.pageContext.list;
    const currentUrl = window.location.href.toLowerCase();

    let listSettingsHtml = '';

    const isListPage =
      currentUrl.indexOf('/lists/') !== -1 ||
      currentUrl.indexOf('/forms/') !== -1 ||
      currentUrl.indexOf('allitems.aspx') !== -1 ||
      currentUrl.indexOf('dispform.aspx') !== -1 ||
      currentUrl.indexOf('editform.aspx') !== -1 ||
      currentUrl.indexOf('newform.aspx') !== -1;

    if (listContext && isListPage) {

      const listId = listContext.id.toString();

      const listSettingsUrl = `${webUrl}/_layouts/15/listedit.aspx?List=${listId}`;

      listSettingsHtml = `
    <li class="${styles.gearMenuItem}">
      <a class="${styles.gearMenuLink}"
         href="${listSettingsUrl}"
         target="_self">
        List / Library Settings
      </a>
    </li>
  `;
    }

    /* ===== FETCH MENU DATA ===== */
    const buItems = await MenuService.getMenuItems(
      'BU',
      this.context.spHttpClient
    );

    const entityItems = await MenuService.getMenuItems(
      'Entity',
      this.context.spHttpClient
    );

    const corporateItems = await MenuService.getMenuItems(
      'Corporate Procedures',
      this.context.spHttpClient
    );

    const appLinksItems = await MenuService.getMenuItems(
      'AppLinks',
      this.context.spHttpClient
    );
    /* ===== BUILD HTML ===== */
    const buHtml = this._renderMenuHtml(
      this._buildMenuTree(buItems, 'BU')
    );

    const entityHtml = this._renderMenuHtml(
      this._buildMenuTree(entityItems, 'Entity')
    );

    const corporateHtml = this._renderMenuHtml(
      this._buildMenuTree(corporateItems, 'Corporate Procedures')
    );

    const appLinksHtml = this._renderMenuHtml(
      this._buildMenuTree(appLinksItems, 'AppLinks')
    );

    /* ================= HEADER HTML ================= */
    this._top.domElement.innerHTML = `
      <div class="${styles.topNav}">

        <!-- Logo -->
         <div>
          <a href="https://bharatpetroleum.sharepoint.com/sites/qa-iconnect-final"
             target="_blank"
             data-interception="off"
             class="${styles.logo}" style="text-decoration: none; color: inherit;">
            
            <img src="https://bharatpetroleum.sharepoint.com/sites/qa-corporate-publishing-hub/SiteAssets/Masterlogo/iconnectlogo.jpeg" alt="iConnect Logo" />
            
             </a>
        </div>
      <!-- Mobile Hamburger -->
  <input type="checkbox" id="navToggle" class="${styles.navToggle}" />
  <label for="navToggle" class="${styles.hamburger}">
    ☰
  </label>
        <ul class="${styles.menu}">

          <!-- Business Units -->
          <li class="${styles.menuItem} ${styles.dropdown}">
            <a class="${styles.link}" href="#" tabindex="0">
              <i class="bi bi-building-fill"></i> Business Units
              <i class="bi bi-caret-down-fill ${styles.dropdownIcon}"></i>
            </a>
            <ul class="${styles.submenu}">
              ${buHtml}
            </ul>
          </li>

          <!-- Entities -->
          <li class="${styles.menuItem} ${styles.dropdown}">
            <a class="${styles.link}" href="#" tabindex="0">
              <i class="bi bi-stack"></i> Entities
              <i class="bi bi-caret-down-fill ${styles.dropdownIcon}"></i>
            </a>
            <ul class="${styles.submenu}">
              ${entityHtml}
            </ul>
          </li>

          <!-- Corporate Procedures -->
          <li class="${styles.menuItem} ${styles.dropdown}">
            <a class="${styles.link}" href="#" tabindex="0">
              <i class="bi bi-file-earmark-post"></i> Corporate Policies
              <i class="bi bi-caret-down-fill ${styles.dropdownIcon}"></i>
            </a>
            <ul class="${styles.submenu}">
              ${corporateHtml}
            </ul>
          </li>

          <!-- STATIC ITEMS -->
          <li class="${styles.menuItem}" >
            <a class="${styles.link}"
            href="https://bharatpetroleum.sharepoint.com/sites/qa-iconnect-final/SitePages/PoliciesAndProcedure.aspx"
            target="_blank"
            data-interception="off">
   
                <i class="bi bi-file-text-fill"></i> SOP & Guidelines
            </a>
          </li>

          <!-- Apps & links -->
          <li class="${styles.menuItem} ${styles.dropdown}">
            <a class="${styles.link}" href="#" tabindex="0">
              <i class="bi bi-folder-symlink-fill"></i> Applications & Links
              <i class="bi bi-caret-down-fill ${styles.dropdownIcon}"></i>
            </a>
            <ul class="${styles.submenu}">
              ${appLinksHtml}
            </ul>
          </li>

          

          <!-- USER MENU -->
          <li class="${styles.menuItem} ${styles.rightMenu}">
            <a class="${styles.link}" href="#" tabindex="0">
              <i class="bi bi-gear-fill"></i>
            </a>
            <ul class="${styles.gearMenu}">
              
             ${listSettingsHtml}
              ${hasAdminAccess ? `
           <li class="${styles.gearMenuItem}">
             <a class="${styles.gearMenuLink}" href="${settingsUrl}" target="_self" tabindex="0"> Settings</a>
           </li>

           <li class="${styles.gearMenuItem}"> <a class="${styles.gearMenuLink}" href="${siteContentsUrl}" target="_self" tabindex="0"> Site Contents </a>
           </li>
           
           ` : ''}
           

              <li class="${styles.gearMenuItem}">
               <a class="${styles.gearMenuLink}" href="#" target="_blank"  data-interception="off" tabindex="0">Help</a>
              </li>
              <li class="${styles.gearMenuItem}">
               <a class="${styles.gearMenuLink}" href="#" target="_blank"  data-interception="off" tabindex="0">Feedback</a>
              </li>
            </ul>
          </li>

        </ul>
      </div>
    `;
  }

  /* ================= FOOTER ================= */




  private _renderBottom(): void {

    if (document.getElementById("bpcl-footer")) return;

    const pageCanvas = document.querySelector("#spPageCanvasContent");

    if (!pageCanvas) return;

    const year = new Date().getFullYear();

    const footerHtml = `
    <div class="${styles.footerBar}">
      <div class="${styles.footerLeft}">
        © ${year} Bharat Petroleum Corporation Limited. All Rights Reserved
      </div>
      <div class="${styles.footerRight}">
        
        <a href="#">Feedback</a>
        <a href="#">Help</a>
      </div>
    </div>
  `;

    const footerContainer = document.createElement("div");
    footerContainer.id = "bpcl-footer";
    footerContainer.innerHTML = footerHtml;

    pageCanvas.appendChild(footerContainer);
  }
}