import { ISPList } from './HelloWorldWebPart';

export default class MockHttpClient  
{
    private static _items: ISPList[] = [
        {Title: 'Mock List', Description : 'This is the description for Mock list.'},
        {Title: 'Mock List 2', Description : 'This is the description for Mock 2 list.'},
        {Title: 'Mock List 3',Description : 'This is the description for Mock list.'}
    ];

    public static get(): Promise<ISPList[]> 
    {
        return new Promise<ISPList[]>((resolve) => {resolve(MockHttpClient._items); });
    }
}