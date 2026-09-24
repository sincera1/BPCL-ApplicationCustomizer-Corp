import { SPHttpClient } from '@microsoft/sp-http';

interface IMenuMasterListItemResponse {
  Id: number;
  Title: string;
  Category: string;
  Level: number;
  Created: string;
  SiteURL?: { Url: string };
  ParentID?: { Id: number };
}

export interface IMenuItem {
  Id: number;
  Title: string;
  Category: string;
  Level: number;
  Created: string;
  SiteURL?: { Url: string };
  ParentIDId?: number;
}

interface ITeamSiteMasterItemResponse {
  Id: number;
  TeamName: string;
  SiteURL?: {
    Url: string;
    Description?: string;
  };
}


export default class MenuService {


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
      `?$select=Id,Title,Category,Level,Created,SiteURL,ParentID/Id` +
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
      Created: item.Created,  
      SiteURL: item.SiteURL,
      ParentIDId: item.ParentID?.Id
    }));
  }


  public static async getTeamSiteUrl(
    teamName: string,
    spHttpClient: SPHttpClient
  ): Promise<string> {

    const baseUrl = this.getBaseUrl();

    const url =
      `${baseUrl}/_api/web/lists/getbytitle('TeamSiteMaster')/items` +
      `?$select=Id,TeamName,SiteURL` +
      `&$filter=TeamName eq '${teamName.replace(/'/g, "''")}'`;

    const response = await spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(
        "Something went wrong while getting Team Site information."
      );
    }

    const data: {
      value: ITeamSiteMasterItemResponse[];
    } = await response.json();

    if (!data.value || data.value.length === 0) {
      return "";
    }

    return data.value[0].SiteURL?.Url || "";
  }
}