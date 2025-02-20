/*
Disclaimer:
The script provided here is for educational and informational purposes only. I do not guarantee its functionality, security, or fitness for any particular purpose. Before running the script, please review and test it thoroughly. By using this script, you agree that you do so at your own risk, and I assume no liability for any issues or damages that may arise from its use. Always ensure you understand and authorize the necessary permissions for your Google Doc and YouTube API access.
*/


function onOpen() {
 DocumentApp.getUi()
   .createMenu('YouTube Tools')
   .addItem('Scrape Comments', 'scrapeCommentsToDoc')
   .addToUi();
}
function scrapeCommentsToDoc() {
 var ui = DocumentApp.getUi();
 // Prompt for the YouTube video ID.
 var response = ui.prompt("YouTube Comment Scraper",
                          "Enter the YouTube video ID (the part after 'v=' in the URL):",
                          ui.ButtonSet.OK_CANCEL);
 if (response.getSelectedButton() != ui.Button.OK) {
   ui.alert("Operation cancelled.");
   return;
 }
 var videoId = response.getResponseText().trim();
 if (!videoId) {
   ui.alert("No video ID provided. Exiting.");
   return;
 }
  var doc = DocumentApp.getActiveDocument();
 var body = doc.getBody();
 body.clear();
  // Fetch video details.
 var videoResponse = YouTube.Videos.list("snippet,statistics", {id: videoId});
 if (!videoResponse.items || videoResponse.items.length === 0) {
   ui.alert("No video found for ID: " + videoId);
   return;
 }
 var videoItem = videoResponse.items[0];
 var title = videoItem.snippet.title;
 var channelName = videoItem.snippet.channelTitle;
 var viewCount = videoItem.statistics.viewCount;
 var commentCountStr = videoItem.statistics.commentCount;
 var totalCommentsStat = parseInt(commentCountStr, 10) || 0;
 var videoLink = "https://www.youtube.com/watch?v=" + videoId;
  // Batch processing variables.
 var changeCount = 0;
 var batchThreshold = 500; // Save and reopen after 500 changes.
  // Helper: Append a paragraph, set hyperlink if needed, then check batch.
 // If indent is provided, sets paragraph indent.
 // If linkData is provided as an object {start, end, url}, then it sets the link.
 function processParagraph(text, indent, linkData) {
   var p = body.appendParagraph(text);
   if (indent) {
     p.setIndentStart(indent);
   }
   if (linkData) {
     // Ensure indices are within bounds.
     var txt = p.getText();
     var start = linkData.start;
     var end = linkData.end;
     if (start < txt.length && end < txt.length) {
       p.editAsText().setLinkUrl(start, end, linkData.url);
     } else {
       Logger.log("Warning: Skipping link due to index error. Text length: " + txt.length + ", start: " + start + ", end: " + end);
     }
   }
   changeCount++;
   if (changeCount % batchThreshold === 0) {
     doc.saveAndClose();
     Utilities.sleep(1000); // Allow some time for changes to commit.
     doc = DocumentApp.openById(doc.getId());
     body = doc.getBody();
   }
   return p;
 }
  // Insert header info.
 processParagraph("Video Comments for \"" + title + "\"", null, null)
     .setHeading(DocumentApp.ParagraphHeading.TITLE);
 processParagraph("Video Link: " + videoLink, null, null);
 processParagraph("Channel Name: " + channelName, null, null);
 processParagraph("Total Views: " + viewCount, null, null);
 processParagraph("Total Comments (stat, top-level + replies): " + commentCountStr, null, null);
 processParagraph("---------------------------------------------", null, null);
  Logger.log("Fetched video details for: " + title);
  var nextPageToken = "";
 var processedTopLevel = 0;
 var totalReplies = 0;
  // Loop through top-level comment threads.
 while (true) {
   var commentThreads = YouTube.CommentThreads.list("snippet", {
     videoId: videoId,
     maxResults: 100,
     pageToken: nextPageToken,
     textFormat: "plainText"
   });
  
   nextPageToken = commentThreads.nextPageToken;
  
   for (var i = 0; i < commentThreads.items.length; i++) {
     processedTopLevel++;
     var item = commentThreads.items[i];
     var topComment = item.snippet.topLevelComment;
     var topSnippet = topComment.snippet;
     var commentId = topComment.id;
    
     var commentText = "Comment by " + topSnippet.authorDisplayName + " (" +
                       topSnippet.publishedAt + "): " + topSnippet.textDisplay +
                       " [Likes: " + topSnippet.likeCount + "]";
    
     // Compute indices for the timestamp within the comment text.
     var prefix = "Comment by " + topSnippet.authorDisplayName + " (";
     var startIndex = prefix.length;
     var timestamp = topSnippet.publishedAt;
     var endIndex = startIndex + timestamp.length - 1;
     var commentLink = "https://www.youtube.com/watch?v=" + videoId + "&lc=" + commentId;
    
     processParagraph(commentText, null, {start: startIndex, end: endIndex, url: commentLink});
     Logger.log("Scraped top-level comment " + processedTopLevel);
    
     // Process replies if available.
     if (item.snippet.totalReplyCount > 0) {
       var parentId = topComment.id;
       var nextPageTokenRep = "";
       processParagraph("Replies:", 30, null);
       do {
         var repliesData = YouTube.Comments.list("snippet", {
           parentId: parentId,
           maxResults: 100,
           pageToken: nextPageTokenRep,
           textFormat: "plainText"
         });
         nextPageTokenRep = repliesData.nextPageToken;
        
         for (var j = 0; j < repliesData.items.length; j++) {
           totalReplies++;
           var replyItem = repliesData.items[j];
           var replySnippet = replyItem.snippet;
           var replyText = "Reply by " + replySnippet.authorDisplayName + " (" +
                           replySnippet.publishedAt + "): " + replySnippet.textDisplay +
                           " [Likes: " + replySnippet.likeCount + "]";
           var rPrefix = "Reply by " + replySnippet.authorDisplayName + " (";
           var rStartIndex = rPrefix.length;
           var rTimestamp = replySnippet.publishedAt;
           var rEndIndex = rStartIndex + rTimestamp.length - 1;
           var replyLink = "https://www.youtube.com/watch?v=" + videoId + "&lc=" + replyItem.id;
           processParagraph(replyText, 50, {start: rStartIndex, end: rEndIndex, url: replyLink});
           Logger.log("   Processed reply " + totalReplies);
         }
       } while (nextPageTokenRep);
     }
     processParagraph("------------------------------------------", null, null);
   }
  
   if (!nextPageToken) {
     break;
   }
 }
  // Insert summary after header.
 var summaryText = "Summary: Video statistics show " + totalCommentsStat + " total comments (top-level + replies). " +
                   "This document contains " + processedTopLevel + " top-level comments and " + totalReplies + " replies.";
 body.insertParagraph(5, summaryText);
  ui.alert("Finished scraping comments and replies.\nTop-level comments: " + processedTopLevel + "\nReplies: " + totalReplies);
}
/*
Disclaimer:
The script provided here is for educational and informational purposes only. I do not guarantee its functionality, security, or fitness for any particular purpose. Before running the script, please review and test it thoroughly. By using this script, you agree that you do so at your own risk, and I assume no liability for any issues or damages that may arise from its use. Always ensure you understand and authorize the necessary permissions for your Google Doc and YouTube API access.
*/
