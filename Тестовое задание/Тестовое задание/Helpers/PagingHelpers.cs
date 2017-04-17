using System;
using System.Text;
using System.Web.Mvc;
using Тестовое_задание.Models;

namespace Тестовое_задание.Helpers
{
    public static class PagingHelpers
    {
        public static MvcHtmlString PageLinks(this HtmlHelper html, 
            PageInfo pageInfo, Func<int, string> pageUrl)
        {
            var startIndex = pageInfo.PageNumber - 5 > 0 ? pageInfo.PageNumber - 5 : 1;
            var endIndex = pageInfo.PageNumber + 5 <= pageInfo.TotalPages ? pageInfo.PageNumber + 5 : pageInfo.TotalPages;

            TagBuilder tagNav = new TagBuilder("nav");
            TagBuilder tagDiv = new TagBuilder("div");
            tagDiv.Attributes.Add("style", "text-align: center;");
            TagBuilder tagUl = new TagBuilder("ul");
            tagUl.AddCssClass("pagination");

            TagBuilder tagLiPrev = new TagBuilder("li");
            TagBuilder tagAPrev = new TagBuilder("a");
            tagAPrev.MergeAttribute("href", pageUrl(pageInfo.PageNumber - 1 > 0 ? pageInfo.PageNumber - 1 : 1));
            tagAPrev.InnerHtml = "&laquo";
            tagLiPrev.InnerHtml += tagAPrev.ToString();
            tagUl.InnerHtml += tagLiPrev.ToString();

            for (int i = startIndex; i <= endIndex; i++)
            {
                TagBuilder tagLi = new TagBuilder("li");
                TagBuilder tagA = new TagBuilder("a");
                
                tagA.MergeAttribute("href", pageUrl(i));
                tagA.InnerHtml = i.ToString();
                if (i == pageInfo.PageNumber)
                {
                    tagLi.AddCssClass("active");
                }
                tagLi.InnerHtml += tagA.ToString();
                tagUl.InnerHtml += tagLi.ToString();
            }

            TagBuilder tagLiNext = new TagBuilder("li");
            TagBuilder tagANext = new TagBuilder("a");
            tagANext.MergeAttribute("href", pageUrl(pageInfo.PageNumber + 1 <= pageInfo.TotalPages ? pageInfo.PageNumber + 1 : pageInfo.TotalPages));
            tagANext.InnerHtml = "&raquo";
            tagLiNext.InnerHtml += tagANext.ToString();
            tagUl.InnerHtml += tagLiNext.ToString();

            tagDiv.InnerHtml += tagUl.ToString();
            tagNav.InnerHtml += tagDiv.ToString();
            return MvcHtmlString.Create(tagNav.ToString());
        }
    }
}