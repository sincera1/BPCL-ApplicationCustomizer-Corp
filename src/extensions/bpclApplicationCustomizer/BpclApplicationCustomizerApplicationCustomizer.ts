import { Log } from '@microsoft/sp-core-library';
import styles from './BpclApplicationCustomizer.module.scss';
import MenuService, { IMenuItem } from './Service';
import { SPPermission } from '@microsoft/sp-page-context';
import AskDiaImage from './assets/AskDia.png';

import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';

interface IMenuNode {
  Id: number;
  Title: string;
  Created: string;
  SiteURL?: { Url: string };
  isNew: boolean;
  children: IMenuNode[];
}

export default class ApplicationCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {

  private _top: PlaceholderContent | undefined;

  private _currentUserSBU: string = "";
  private _sbuPromise: Promise<string> | undefined;



  public async onInit(): Promise<void> {

    Log.info('ApplicationCustomizer', 'Initialized');

    this._hideAppBar();
    this._loadBootstrapIcons();

    await this._renderTop();

    this._renderBottom();

    this._sbuPromise = this.getCurrentUserSBUFromGraph();

    this._sbuPromise.then((sbu) => {
      this._currentUserSBU = sbu;
    });

    this.context.application.navigatedEvent.add(this, async () => {

      // Recreate the top placeholder so menu visibility updates
      if (this._top) {
        this._top.dispose();
        this._top = undefined;
      }

      await this._renderTop();

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

  public async getCurrentUserSBUFromGraph(): Promise<string> {
    try {
      const client = await this.context.msGraphClientFactory.getClient("3");

      const user = await client
        .api("/me?$select=onPremisesExtensionAttributes")
        .get();

      return (
        user?.onPremisesExtensionAttributes
          ?.extensionAttribute4
          ?.trim() || ""
      );

    } catch (error) {
      console.error(
        "ApplicationCustomizer.getCurrentUserSBUFromGraph",
        error
      );

      return "";
    }
  }

  /* ================= HIDE SHAREPOINT CHROME ================= */
  private _hideAppBar(): void {

    const style = document.createElement('style');
    style.innerHTML = `
      #spSiteHeader { display: none !important; }
      #sp-appBar { display: none !important; }
      .ms-HorizontalNav { margin-left: 0 !important; }
     
     #CommentsWrapper {display: none !important;}
      
    #spCommandBar { 
        display: none !important; 
      }
    
      #vpc_Page\\.SiteFooter\\.internal\\.03025612-a400-4804-a78e-e1493200a43b { display: none !important; }
      #CommentsWrapper { display: none !important; }
      #O365_MainLink_Settings {
      display: none !important;
    }

       #spLeftNav {
      display: none !important;
    }

     
    [data-automationid="appHeaderBar"] {
      display: none !important;
    }
    
    `;
    document.head.appendChild(style);
  }





  /* ================= BUILD LEVEL 0 → LEVEL 1 → LEVEL 2 TREE ================= */

  private _buildMenuTree(
    items: IMenuItem[],
    category: string
  ): IMenuNode[] {

    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

    const isNewItem = (created: string): boolean => {
      if (!created) {
        return false;
      }

      const createdDate = new Date(created);

      return createdDate >= sevenDaysAgo;
    };

    const parents = items.filter(
      item =>
        item.Level === 0 &&
        item.Category === category
    );

    return parents.map(parent => {

      const parentChildren = items
        .filter(
          child =>
            child.Level === 1 &&
            child.ParentIDId === parent.Id
        )
        .map(child => {

          const childSubChildren = items
            .filter(
              subChild =>
                subChild.Level === 2 &&
                subChild.ParentIDId === child.Id
            )
            .map(subChild => ({
              Id: subChild.Id,
              Title: subChild.Title,
              SiteURL: subChild.SiteURL,
              Created: subChild.Created,
              isNew: isNewItem(subChild.Created),
              children: []
            }));

          return {
            Id: child.Id,
            Title: child.Title,
            SiteURL: child.SiteURL,
            Created: child.Created,
            isNew:
              isNewItem(child.Created) ||
              childSubChildren.some(subChild => subChild.isNew),
            children: childSubChildren
          };
        });

      return {
        Id: parent.Id,
        Title: parent.Title,
        SiteURL: parent.SiteURL,
        Created: parent.Created,
        isNew:
          isNewItem(parent.Created) ||
          parentChildren.some(child => child.isNew),
        children: parentChildren
      };
    });
  }

  /* ================= RENDER DROPDOWN HTML ================= */



  private _renderMenuHtml(menuTree: IMenuNode[]): string {

    const renderNewBadge = (isNew: boolean): string => {
      return isNew
        ? `<span class="${styles.newBadge}">NEW</span>`
        : '';
    };

    const renderItems = (items: IMenuNode[]) => items.map(parent => `

    <li class="${styles.submenuItem}">

      ${parent.children.length > 0
        ? `
          <a class="${styles.submenuLink} d-flex align-items-center justify-content-between">
          <span class="d-flex align-items-center">
            <span>

              ${parent.Title}
              ${renderNewBadge(parent.isNew)}
            </span>
            </span>

            <i class="bi bi-caret-right-fill"></i>
          </a>
        `
        : `
          <a class="${styles.submenuLink}"
             href="${parent.SiteURL?.Url || '#'}"
             target="_blank"
             data-interception="off">

            ${parent.Title}
            ${renderNewBadge(parent.isNew)}

          </a>
        `
      }

      ${parent.children.length > 0
        ? `
          <ul class="${styles.rightSubmenu}">

            ${parent.children.map(child => `

              <li class="${styles.submenuItem}">

                ${child.children.length > 0
            ? `
                    <a class="${styles.submenuLink} d-flex align-items-center justify-content-between">
                    <span class="d-flex align-items-center">
                      <span>
                        ${child.Title}
                        ${renderNewBadge(child.isNew)}
                      </span>
                    </span>
                      <i class="bi bi-caret-right-fill"></i>

                    </a>
                  `
            : `
                    <a class="${styles.submenuLink}"
                       href="${child.SiteURL?.Url || '#'}"
                       target="_blank"
                       data-interception="off">

                      ${child.Title}
                      ${renderNewBadge(child.isNew)}

                    </a>
                  `
          }

                ${child.children.length > 0
            ? `
                    <ul class="${styles.rightSubmenu}">

                      ${child.children.map(subChild => `

                        <li>

                          <a href="${subChild.SiteURL?.Url || '#'}"
                             target="_blank"
                             data-interception="off">

                            ${subChild.Title}
                            ${renderNewBadge(subChild.isNew)}

                          </a>

                        </li>

                      `).join('')}

                    </ul>
                  `
            : ''
          }

              </li>

            `).join('')}

          </ul>
        `
        : ''
      }

    </li>

  `).join('');


    // 12 items per column
    const itemsPerColumn = 12;

    const columns: IMenuNode[][] = [];

    for (let i = 0; i < menuTree.length; i += itemsPerColumn) {

      columns.push(
        menuTree.slice(i, i + itemsPerColumn)
      );

    }


    // If only one column
    if (columns.length === 1) {

      return renderItems(menuTree);

    }


    return `
    <li class="${styles.submenuItem} ${styles.submenuColumnsWrapper}">

      <div class="${styles.submenuWrapper}">

        ${columns.map(col => `

          <ul class="${styles.column}">

            ${renderItems(col)}

          </ul>

        `).join('')}

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

    const canEditPage =
      this.context.pageContext.web.permissions.hasPermission(
        SPPermission.manageWeb
      );

    const currentUrl = window.location.href.toLowerCase();
    const siteUrl = this.context.pageContext.web.absoluteUrl.toLowerCase();

    const isDashboard =


      // Dashboard as home page
      currentUrl === siteUrl ||
      currentUrl === siteUrl + "/" ||
      currentUrl.indexOf("/sitepages/dashboard.aspx") !== -1 ||
      currentUrl.indexOf("/sitepages/viewallnews.aspx") !== -1 ||
      currentUrl.indexOf("/sitepages/viewallbroadcast.aspx") !== -1 ||
      currentUrl.indexOf("/sitepages/viewallevents.aspx") !== -1;

    const isSitePage =
      currentUrl.indexOf("/sitepages/") > -1;

    const showEditPage =
      canEditPage &&
      isSitePage &&
      !isDashboard;

    /* ===== LIST / LIBRARY DETECTION ===== */
    const listContext = this.context.pageContext.list;


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

    const isNewWithin7Days = (created: string): boolean => {
      if (!created) {
        return false;
      }

      const createdDate = new Date(created);
      const sevenDaysAgo = new Date();

      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

      return createdDate >= sevenDaysAgo;
    };

    const hasNewBU = buItems.some(item =>
      isNewWithin7Days(item.Created)
    );

    const hasNewEntity = entityItems.some(item =>
      isNewWithin7Days(item.Created)
    );

    const hasNewCorporate = corporateItems.some(item =>
      isNewWithin7Days(item.Created)
    );

    const hasNewAppLinks = appLinksItems.some(item =>
      isNewWithin7Days(item.Created)
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
          <a href="https://bharatpetroleum.sharepoint.com/sites/dev-iconnect-final"
             target="_blank"
             data-interception="off"
             class="${styles.logo}" style="text-decoration: none; color: inherit;">
            
            <img src="https://bharatpetroleum.sharepoint.com/sites/dev-corporate-publishing-hub/SiteAssets/Masterlogo/iconnectlogo.jpeg" alt="iConnect Logo" />
            
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
              ${hasNewBU ? `<span class="${styles.newBadge}">NEW</span>` : ''}
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
              ${hasNewEntity ? `<span class="${styles.newBadge}">NEW</span>` : ''}
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
              ${hasNewCorporate ? `<span class="${styles.newBadge}">NEW</span>` : ''}
              <i class="bi bi-caret-down-fill ${styles.dropdownIcon}"></i>
            </a>
            <ul class="${styles.submenu}">
              ${corporateHtml}
            </ul>
          </li>

          <!-- STATIC ITEMS -->
          <li class="${styles.menuItem}" >
            <a class="${styles.link}"
            href="https://bharatpetroleum.sharepoint.com/sites/iconnect/SitePages/PoliciesAndProcedure.aspx"
            target="_blank"
            data-interception="off">
   
                <i class="bi bi-file-text-fill"></i> SOP & Guidelines
            </a>
          </li>

          <!-- Apps & links -->
          <li class="${styles.menuItem} ${styles.dropdown}">
            <a class="${styles.link}" href="#" tabindex="0">
              <i class="bi bi-folder-symlink-fill"></i> Applications & Links
              ${hasNewAppLinks ? `<span class="${styles.newBadge}">NEW</span>` : ''}
              <i class="bi bi-caret-down-fill ${styles.dropdownIcon}"></i>
            </a>
            <ul class="${styles.submenu}">
              ${appLinksHtml}
            </ul>
          </li>

          <!-- My Team -->
        <li class="${styles.menuItem}">
       <a  class="${styles.link}"  href="#"  tabindex="0" id="myTeamMenu">
  
        <i class="bi bi-people-fill"></i> My Team
        </a>
        </li>
        

          
     <li class="${styles.menuItem} ${styles.askDiaMenu}">
<a
        href="https://dia.bpcl.in/"
        target="_blank"
        data-interception="off"
        class="${styles.askDiaLink}"
>
<div class="${styles.askDiaContainer}">
<div class="${styles.askDiaIcon}">
<img
                    src="${AskDiaImage}"
                    alt="Ask DIA"
                />
</div>
 
            <span class="${styles.askDiaText}">
                Ask DIA
</span>
</div>
</a>
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
           ${showEditPage ? `
            <li class="${styles.gearMenuItem}"><a id="customEditPage" class="${styles.gearMenuLink}"href="javascript:void(0)">Edit Page </a>
            </li>
            ` : ''}
       
  
              <li class="${styles.gearMenuItem}">
               <a class="${styles.gearMenuLink}" href="https://bharatpetroleum.sharepoint.com/sites/iconnect/SitePages/User_Guide.aspx" target="_blank"  data-interception="off" tabindex="0">Help</a>
              </li>
              <li class="${styles.gearMenuItem}">
               <a class="${styles.gearMenuLink}" href="https://bharatpetroleum.sharepoint.com/sites/iconnect/SitePages/Feedback.aspx" target="_blank"  data-interception="off" tabindex="0">Feedback</a>
              </li>
            </ul>
          </li>

        </ul>
      </div>
    `;

    const myTeamMenu = document.getElementById("myTeamMenu");

    if (myTeamMenu) {
      myTeamMenu.addEventListener("click", async (event) => {
        event.preventDefault();

        if (!this._currentUserSBU && this._sbuPromise) {
          this._currentUserSBU = await this._sbuPromise;
        }

        const sbu = this._currentUserSBU;

        // If SBU is not available, do nothing
        if (!sbu) {
          return;
        }

        // My Team logic will continue here
      });
    }
    const editPageBtn = document.getElementById('customEditPage');


    if (editPageBtn) {

      editPageBtn.addEventListener('click', () => {


        // remove css hiding
        const style = document.createElement('style');

        style.innerHTML = `
      #spCommandBar {
        display: flex !important;
      }
    `;

        document.head.appendChild(style);


        setTimeout(() => {

          const editButton = document.querySelector(
            '[aria-label="Edit"]'
          ) as HTMLElement;

          editButton?.click();

        }, 300);

      });

    }


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
        
        <a href="https://bharatpetroleum.sharepoint.com/sites/dev-iconnect-final/SitePages/Feedback.aspx"
        target="_blank"
        data-interception="off"
        rel="noopener noreferrer">
        Feedback
        </a>
        <a href="https://bharatpetroleum.sharepoint.com/sites/dev-iconnect-final/SitePages/User_Guide.aspx"
        target="_blank"
        data-interception="off"
        rel="noopener noreferrer"
        >Help</a>
      </div>
    </div>
  `;

    const footerContainer = document.createElement("div");
    footerContainer.id = "bpcl-footer";
    footerContainer.innerHTML = footerHtml;

    pageCanvas.appendChild(footerContainer);
  }
}