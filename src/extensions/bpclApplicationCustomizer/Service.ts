import { SPHttpClient } from '@microsoft/sp-http';

interface IMenuMasterListItemResponse {
  Id: number;
  Title: string;
  Category: string;
  Level: number;
  SiteURL?: { Url: string };
  ParentID?: { Id: number };
}

export interface IMenuItem {
  Id: number;
  Title: string;
  Category: string;
  Level: number;
  SiteURL?: { Url: string };
  ParentIDId?: number;
}

export default class MenuService {

  // private static BASE_URL =
  //   'https://bharatpetroleum.sharepoint.com/sites/qa-corporate-publishing-hub/';

  //  Fixed site path (no tenant hardcoding)
  private static HUB_SITE = "/sites/dev-corporate-publishing-hub";

  //  Dynamic base URL (tenant comes automatically)
  private static getBaseUrl(): string {
    return `${window.location.origin}${this.HUB_SITE}`;
  }

  public static async getMenuItems(
    category: string,
    spHttpClient: SPHttpClient
  ): Promise<IMenuItem[]> {

    const baseUrl = this.getBaseUrl();

    const url =
      `${baseUrl}/_api/web/lists/getbytitle('MenuMasterList')/items` +
      `?$select=Id,Title,Category,Level,SiteURL,ParentID/Id` +
      `&$expand=ParentID` +
      `&$filter=Category eq '${category}'` +
      `&$orderby=Level,Title`;

    const response = await spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Something went wrong. Please contact administrator.`);
    }

    const data: { value: IMenuMasterListItemResponse[] } = await response.json();

    // 🔁 Normalize lookup Id
    return data.value.map((item) => ({
      Id: item.Id,
      Title: item.Title,
      Category: item.Category,
      Level: item.Level,
      SiteURL: item.SiteURL,
      ParentIDId: item.ParentID?.Id
    }));
  }
}