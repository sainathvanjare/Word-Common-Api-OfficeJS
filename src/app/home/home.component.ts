import { Component, OnInit } from '@angular/core';
import { AppService } from "../app.service";
import * as $ from 'jquery';
import { range } from 'rxjs';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss']
})
export class HomeComponent implements OnInit {

  
  constructor(private appService: AppService) {
  }
  title = 'word-addin';
  selectedStyle: any;
  textStyles = [
    { name: "Spartan Styles", value: "" },
    { name: "Heading1", value: "Heading1" },
    { name: "Heading2", value: "Heading2" },
    { name: "Heading3", value: "Heading3" },
    { name: "Bullet Text", value: "Bullet Text" }
    // { name: "Body Text", value: "Body Text" }
  ]

  ngOnInit() {
    // this.autoShowTaskpane();
    // this.initData();

    // this.officeExe();
    // this.getDocumentData();
  }

  async officeExe(data) {
    // Gets the current selection and changes the font color to red.
    await Word.run(async (context: any) => {
      // this.getBookmarkRange(context, "bkCurrentPrice")
      this.getBkAndReplaceWithValues(context, data);
      let range = context.document.body.getRange();
      await context.sync();
      // const bkRange = await this.getBookmarkRange(context, 'test2')
      // this.changeStyle(context, range, Word.Style.heading1)
      // await context.sync()
      // bkRange.insertText('Test text ', 'Replace')
      // await context.sync()
    });
  }


  initData() {
    let docId = this.getCurrentDocumentName();
    this.appService.getDataByDocumentId(docId).subscribe((data) => {
      console.log(data.success)
      this.officeExe(data)
    })

  }
  async printSelection() {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.font.color = "red";
      range.load("text");
      await context.sync();
    });
  }

  
async insertTable1() {
  await Word.run(async (context) => {
    // Use a two-dimensional array to hold the initial table values.
    let data = [
      ["Tokyo", "Beijing", "Seattle"],
      ["Apple", "Orange", "Pineapple"]
    ];
    let table = context.document.body.insertTable(2, 3, "Start", data);
    table.styleBuiltIn = Word.Style.gridTable5Dark_Accent2;
    table.styleFirstColumn = false;

    await context.sync();
  });
}

async  insertComment() {
  await Word.run(async (context: any) => {
    let text = $("#This is test comment")
      .val()
      .toString();
    let comment = context.document.getSelection().insertComment(text);

    // Load object for display in Script Lab console.
    comment.load();
    await context.sync();

    console.log("Comment inserted:");
    console.log(comment);
  });
}

  async changeTextColor(context, range, color) {
    range.font.color = color;
    await context.sync()
  }

  async changeStyle(context, range, style) {
    // Word.Style.heading1
    range.styleBuiltIn = style;
    await context.sync()
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

  autoShowTaskpane() {
    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();
  }

  getCurrentDocumentName() { 
    let docUrl = Office.context.document.url;
    let docName = "";
    if (docUrl) {
      let arr: any = "";
      if (docUrl.indexOf("/") > -1) {
        arr = docUrl.toString().split("/");
      }
      else {
        arr = docUrl.toString().split("\\");
      }

      docName = arr[arr.length - 1];
      docName = docName.substring(0, docName.lastIndexOf('.'));
    }
    return docName;
  }
  async getBkAndReplaceWithValues(context, data) {
    // let bookmarkList: any = this.getAllBookmarks(Office.context, range);
    // bookmarkList.forEach(bookmark => {
    //   if (bookmark.text == "bkCurrentPrice") {

    //   }
    // });

    this.insertTable(context, "");
    this.replaceBookMarkText(context, "bkCurrentPrice", "$999.99");
    this.replaceBookMarkText(context, "bkAnalystName1", data.actionNote.metadata.primaryAuthor[0].displayName);
    this.replaceBookMarkText(context, "bkAnalystPhone1", data.actionNote.metadata.primaryAuthor[0].workPhone);
    this.replaceBookMarkText(context, "bkCompanyName", data.actionNote.primaryIssuer.issuer.name);
    this.replaceBookMarkText(context, "bkCompanyTicker", data.actionNote.primaryIssuer.issuer.security.ticker[0].name);
    this.replaceBookMarkText(context, "bkCompanyRecommendation", data.actionNote.primaryIssuer.recommendation);
    this.replaceBookMarkText(context, "bkTargetPrice", data.actionNote.primaryIssuer.targetPrice);

    // bkRange.load();
    await context.sync();
  }

  async replaceBookMarkText(context, bookmarkName, text) {
    let bookmarkRange = context.document.getBookmarkRange(bookmarkName);
    bookmarkRange.load();
    bookmarkRange.insertText(text, 'Replace')
    // await context.sync();
  }

  // getDocumentData() {
  //   this.appService.getDataByDocumentId().subscribe((data: any[]) => {
  //     console.log(data);
  //   })
  // }

  async insertTable(context, bookmarkName) {
    let data = [
      ["Company Financials"],
      ["Tokyo", "Beijing", "Seattle"],
      ["Pune", "Mumbai", "Delhi"],
      ["Apple", "Orange", "Pineapple"],
      [""]
    ];
    let bookmarkRange = context.document.getBookmarkRange("bkRatingTable1");
    bookmarkRange.load();
    let table = bookmarkRange.insertTable(5, 3, "Before", data);
    table.mergeCells(0, 0, 0, 2); // [start row num, no of row to merge, end row no, no of column to merge]
    table.mergeCells(4, 0, 4, 2);
    // this.deleteBookmark(context, "bkRatingTable1")
    // table.styleBuiltIn = Word.Style.gridTable5Dark_Accent2;
    // table.styleFirstColumn = false;
    // await context.sync();
  }

  async addHeader() {
    await Word.run(async (context) => {
      context.document.sections
        .getFirst()
        .getHeader("Primary")
        .insertParagraph("This is a header", "End");

      await context.sync();
    });
  }

  async addFooter() {
    await Word.run(async (context) => {
      context.document.sections
        .getFirst()
        .getFooter("Primary")
        .insertParagraph("This is a footer", "End");

      await context.sync();
    });
  }

  async applyStyle() {
    if (this.selectedStyle.name == "Bullet Text") {
      this.applyBulletStyle();
    }
    else {
      await Word.run(async (context: any) => {
        var range = context.document.getSelection();
        range.load();
        this.changeStyle(context, range, this.selectedStyle.name)

      });
    }
  }


  async insertPage(context) {
    var body = context.document.body;
    // Queue a command to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    await context.sync();
  }

  async applyBulletStyle() {
    await Word.run(async (context) => {
      var paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load();
      await context.sync();
      paragraphs.items.forEach((item) => {
        if (item.text != "") {
          var list = item.startNewList();
          list.load("$none")
          list.setLevelBullet(1, Word.ListBullet.solid)
        }
      })
      await context.sync();
    })
  }
}

