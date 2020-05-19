import { Injectable } from '@angular/core';
import { Observable, of } from 'rxjs';
import { HttpClient, HttpHeaders } from '@angular/common/http';

@Injectable({ providedIn: 'root' })
export class AppService {


    constructor(private http: HttpClient) {

    }

    private api_url = "https://scs.jovus.com/spartan-api/api/documents/fetchdocument";

    httpOptions = {
        headers: new HttpHeaders({ 'Content-Type': 'application/json' })
    };
    getDataByDocumentId(docId): Observable<any> {
        // const url = `${this.api_url}/?documentid=AR_CHTR_01312020`
        const url = `${this.api_url}/?documentid=${docId}`
        return this.http.post(url, {})
    }

    async getAllBookmarks(context, range) {
        const bookmarks = range.getBookmarks();
        await context.sync();
        return bookmarks.value
    }

    async insertBookmark(context, range, bookmarkName) {
        range.insertBookmark(bookmarkName);
        await context.sync();
    }

    async getBookmarkRange(context, bookmarkName) {
        const bookmarkRange = context.document.getBookmarkRange(bookmarkName);
        await context.sync();
        return bookmarkRange
    }

    async deleteBookmark(context, bookmarkName) {
        context.document.deleteBookmark(bookmarkName);
        await context.sync();
    }

}
