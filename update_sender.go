package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/streadway/amqp"
	"encoding/json"
	"fmt"
	"regexp"
	"os"
	"time"
	"strings"
	"strconv"
	"golang.org/x/sys/windows/registry"
)

type update_struct struct{
	KB     string  `json:"kb"`
	Tittle string  `json:"tittle"`
	Software string  `json:"software"`
}

type message_struct struct {
	Hostname string `json:"hostname"`
	Time     int64  `json:"time"`
	Updates  []update_struct
}

func panicOnError(err error, msg string) {
	if err != nil {
		elog.Error(1, fmt.Sprintf("%s: %s", msg, err))
		elog.Error(1, "Panicking")
		panic(err)
	}
}

func getUpdatesList(query string)  ([]update_struct, error) {
	nil_slice := make([]update_struct, 0)
	k, err := registry.OpenKey(registry.LOCAL_MACHINE, `SOFTWARE\Microsoft\Windows NT\CurrentVersion`, registry.QUERY_VALUE)
	if err != nil {
		panicOnError(err,"Failed to open registry CurrentVersion")
	}
	ProductName , _, err := k.GetStringValue("ProductName")
	if err != nil {
		panicOnError(err, "Failed to open registry ProductName")
	}
	defer k.Close()
	ole.CoInitialize(0)
	defer ole.CoUninitialize()
	unknown, err := oleutil.CreateObject("Microsoft.Update.Session")
	if err != nil {
		return  nil_slice, fmt.Errorf("Unable to create initial object, %s", err.Error())
	}
	defer unknown.Release()
	update, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil_slice, fmt.Errorf("Unable to create query interface, %s", err.Error())
	}
	defer update.Release()
	oleutil.PutProperty(update, "ClientApplicationID", "GoLang Windows API")

	us, err := oleutil.CallMethod(update, "CreateUpdateSearcher")
	if err != nil {
		return nil_slice, fmt.Errorf("Error creating update searcher, %s", err.Error())
	}
	usd := us.ToIDispatch()
	defer usd.Release()

	usr, err := oleutil.CallMethod(usd, "Search", query)
	if err != nil {
		return nil_slice, fmt.Errorf("Error performing update search, %s", err.Error())
	}
	usrd := usr.ToIDispatch()
	defer usrd.Release()

	if err != nil {
		return nil_slice, fmt.Errorf("Error getting Updates collection, %s", err.Error())
	}
	rules := oleutil.MustGetProperty(usrd, "Updates").ToIDispatch()

	newEnum, err := rules.GetProperty("_NewEnum")
	if err != nil {
		return nil_slice, err
	}
	defer newEnum.Clear()

	enum, err := newEnum.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)
	if err != nil {
		return nil_slice, err
	}
	defer enum.Release()
	updates := make([]update_struct, 0)
	r, _ := regexp.Compile("KB\\w*")
	for item, length, err := enum.Next(1); length > 0; item, length, err = enum.Next(1) {
		if err != nil {
			return nil_slice, err
		}
		rule := item.ToIDispatch()
		name := oleutil.MustGetProperty(rule, "Title").ToString()

		updates = append(updates, update_struct{KB: strings.ToLower(r.FindString(name)), Tittle: name, Software: strings.ToLower(ProductName)})
	}
	return updates, nil
}

func makeMessage() []byte{
	updates_struct, err := getUpdatesList("IsInstalled=0")

	if err != nil {
		fmt.Println(err)
		return nil
	}
	raw_message := new(message_struct)
	hostname, _ := os.Hostname()
	raw_message.Hostname = strings.ToLower(hostname)
	start := time.Now()
	raw_message.Time = start.Unix()
	raw_message.Updates = updates_struct
	messageJson, _ := json.Marshal(raw_message)
	return messageJson
}
func sendtoRabbitMQ(message []byte, config rabbitConfig){
	url := "amqp://" + config.User + ":" + config.Pass + "@" + config.Host + ":" + strconv.Itoa(config.Port) + "/" + config.Vhost
	conn, err := amqp.Dial(url)
	panicOnError(err, "Failed to connect to RabbitMQ")
	defer conn.Close()

	ch, err := conn.Channel()
	panicOnError(err, "Failed to open a channel")
	defer ch.Close()

	q, err := ch.QueueDeclare(
		config.Queue,
		true,
		false,
		false,
		false,
		nil,
	)
	panicOnError(err, "Failed to declare a queue")

	body := message
	err = ch.Publish(
		"win_kb_info",     // exchange
		q.Name, // routing key
		false,  // mandatory
		false,  // immediate
		amqp.Publishing{
			ContentType: "text/plain",
			Body:        []byte(body),
		})
	panicOnError(err, "Failed to publish a message")
}


func updateSender(config rabbitConfig) {
	defer func() {
		if r := recover(); r != nil {
			elog.Info(1, "Recovered")
		}
	}()
	elog.Info(1, "Get update list")
	messageJson := makeMessage()
	elog.Info(1, fmt.Sprintf("Sending: %s ", messageJson))
	sendtoRabbitMQ(messageJson, config)
}
