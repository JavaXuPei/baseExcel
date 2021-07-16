package main

import (
	"baseExcel/controller"
	"log"
	"net/http"
)

func main() {
	http.HandleFunc("/index", controller.GetExcel)
	err := http.ListenAndServe(":9090", nil)
	if err != nil {
		log.Fatal("ListenAndServe error: ", err)
	}
}
