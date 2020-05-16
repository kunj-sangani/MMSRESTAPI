export interface IAllGroups {
    createdDateTime?: string;
    description?: string;
    id: number;
    lastModifiedDateTime: string;
    name: string;
    type: string;
}

export interface IAllSets {
    createdDateTime?: string;
    description?: string;
    id: number;
    isOpen: boolean;
    groupId: string;
    childrenCount: number;
    localizedNames: IlocalizedNames[];
}

export interface IlocalizedNames {
    name: string;
    languageTag: string;
}

export interface IAllTerms {
    id: string;
    isDeprecated: boolean;
    childrenCount: number;
    createdDateTime: string;
    lastModifiedDateTime: string;
    labels: Ilabels[];
    descriptions: Idescriptions[];
    isAvailableForTagging: IisAvailableForTagging[];
}

export interface Ilabels {
    name: string;
    isDefault: boolean;
    languageTag: string;
}

export interface Idescriptions {
    description: string;
    languageTag: string;
}

export interface IisAvailableForTagging {
    setId: string;
    isAvailable: boolean;
}