package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"

	"github.com/PuerkitoBio/goquery"
	"github.com/gocolly/colly"
	"github.com/xuri/excelize/v2"
)

type Hot struct {
	Movie_name string `selector:"div.item>div.info>div.hd>a>span:nth-child(1)"`
	Href       string `selector:"div.item>div.info>div.hd>a[href]"`
	Rating     string `selector:"div.item>div.info>div.bd>div.star>span.rating_num"`
	Playable   string `selector:"div.item>div.info>div.hd>span.playable"`
}

func main() {
	f, err := excelize.OpenFile("./films.xlsx")
	if err != nil && strings.Contains(err.Error(), "no such file") {
		f = WriteFile()
	}
	defer f.Close()
	hots := make([]*Hot, 0)
	c := colly.NewCollector(
		colly.AllowedDomains("movie.douban.com"),
	)
	c.OnResponse(func(r *colly.Response) {
		fmt.Println(r.StatusCode)
	})
	c.OnHTML("ol.grid_view>li", func(h *colly.HTMLElement) {
		hot := &Hot{}

		h.Unmarshal(hot)

		h.DOM.Find("div.item>div.info").Each(func(i int, s *goquery.Selection) {

			title := s.Find("div.hd>a")
			href, _ := title.Attr("href")
			hot.Href = href
			// movie_name := title.Find("span.title").Text()
			// playable := s.Find("div.hd>span.playable").Text()
			// rating := s.Find("div.bd>div.star>span.rating_num").Text()
			// hot.Rating = rating
			// fmt.Println("电影名：", movie_name, "评分：", rating, "地址：", href, playable)

		})

		hots = append(hots, hot)
	})
	c.Visit("https://movie.douban.com/top250")
	fmt.Println("host length:", len(hots))
	for index, val := range hots {
		f.SetSheetRow("Sheet1", "A"+strconv.Itoa(index+2), &[]any{val.Movie_name, val.Rating, val.Href, val.Playable})
	}
	f.Save()
}

func WriteFile() *excelize.File {
	f := excelize.NewFile()
	index := f.NewSheet("Sheet1")
	f.SetSheetRow("Sheet1", "A1", &[]any{"电影名", "评分", "地址", "是否可观看"})
	f.SetRowHeight("Sheet1", 1, 30)       //设置行高度
	f.SetColWidth("Sheet1", "A", "A", 40) //设置列宽度
	f.SetColWidth("Sheet1", "C", "C", 40)
	f.SetColWidth("Sheet1", "D", "D", 12)
	setTitleStyle(f)
	f.SetActiveSheet(index)
	if err := f.SaveAs("films.xlsx"); err != nil {
		log.Fatalln("err", err)
		panic(err)
	}
	return f
}

// 设置标题栏样式
func setTitleStyle(f *excelize.File) {
	A1, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Vertical: "center"}, //垂直居中
		Font:      &excelize.Font{Family: "黑体", Size: 14},
	})
	if err != nil {
		fmt.Println("set style error:", err)
	}
	B1, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "center"}, //水平垂直居中
		Font:      &excelize.Font{Family: "黑体", Size: 14},
	})
	col, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "center"}, //水平垂直居中
	})
	f.SetCellStyle("Sheet1", "A1", "A1", A1)
	f.SetCellStyle("Sheet1", "C1", "C1", A1)
	f.SetColStyle("Sheet1", "B", col) //注意顺序，excelize设置样式会覆盖
	f.SetCellStyle("Sheet1", "B1", "B1", B1)
	f.SetColStyle("Sheet1", "D", col)
	f.SetCellStyle("Sheet1", "D1", "D1", B1)
}
