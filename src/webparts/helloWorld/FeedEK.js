import * as $ from "jquery";
!(function(a) {
  a.fn.FeedEk = function(b) {
    var g,
      c = a.extend(
        {
          FeedUrl: "http://feeds.reuters.com/Reuters/worldNews",
          MaxCount: 5,
          ShowDesc: !0,
          ShowPubDate: !0,
          DescCharacterLimit: 0,
          TitleLinkTarget: "_blank",
          DateFormat: "",
          DateFormatLang: "en"
        },
        b
      ),
      d = a(this).attr("id"),
      f = "";
    a("#" + d)
      .empty()
      .append(
        '<img src="//cdnjs.cloudflare.com/ajax/libs/unitegallery/1.7.28/images/loader-white1.gif" />'
      );
    var h =
      'SELECT channel.item FROM feednormalizer WHERE output="rss_2.0" AND url ="' +
      c.FeedUrl +
      '" LIMIT ' +
      c.MaxCount;
    a.ajax({
      url:
        "https://query.yahooapis.com/v1/public/yql?q=" +
        encodeURIComponent(h) +
        "&format=json&diagnostics=true&callback=?",
      dataType: "json",
      success: function(b) {
        a("#" + d).empty(),
          a.each(b.query.results.rss, function(b, d) {
            if (
              ((f +=
                '<li><div class="itemTitle"><a href="' +
                d.channel.item.link +
                '" target="' +
                c.TitleLinkTarget +
                '" >' +
                d.channel.item.title +
                "</a></div>"),
              c.ShowPubDate)
            ) {
              if (
                ((g = new Date(d.channel.item.pubDate)),
                (f += '<div class="itemDate">'),
                a.trim(c.DateFormat).length > 0)
              )
                try {
                  moment.lang(c.DateFormatLang),
                    (f += moment(g).format(c.DateFormat));
                } catch (b) {
                  f += g.toLocaleDateString();
                }
              else f += g.toLocaleDateString();
              f += "</div>";
            }
            c.ShowDesc &&
              ((f += '<div class="itemContent">'),
              (f +=
                c.DescCharacterLimit > 0 &&
                d.channel.item.description.length > c.DescCharacterLimit
                  ? d.channel.item.description.substring(
                      0,
                      c.DescCharacterLimit
                    ) + "..."
                  : d.channel.item.description),
              (f += "</div>"));
          }),
          a("#" + d).append('<ul class="feedEkList">' + f + "</ul>");
      }
    });
  };
})($);
