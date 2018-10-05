export interface ISpSiteDesignerState {
    siteScriptResults?: any;
    siteDesignResults?: any;
    siteDesignTitle?: string;
    selectedSiteDesignID?: string;
    siteDesignDescription?: string;
    siteDesignWebTemplate?: string;
    siteDesignPreviewImageUrl?: string;
    siteDesignPreviewImageAltText?: string;
    selectedSiteScriptID?: any;
    loading?: boolean;
    siteScriptCharacterCount?: number;
    siteScriptForm?: {
        title?: string;
        content?: any;
        description?: string;
    };
    siteDesignForm?: any;
    siteScriptActionCount: number;
}