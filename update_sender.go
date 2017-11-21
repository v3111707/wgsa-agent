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
	"strconv"
)

type update_struct struct{
	KB     string  `json:"kb"`
	Tittle string  `json:"tittle"`
}

type message_struct struct {
	Hostname string `json:"hostname"`
	Time     int64  `json:"time"`
	Updates  []update_struct
}

func getUpdatesList(query string)  ([]update_struct, error) {
	nil_slice := make([]update_struct, 0)

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
		updates = append(updates, update_struct{KB: r.FindString(name), Tittle: name})
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
	raw_message.Hostname, _ = os.Hostname()
	start := time.Now()
	raw_message.Time = start.Unix()
	raw_message.Updates = updates_struct
	messageJson, _ := json.Marshal(raw_message)
	return messageJson
}
func sendtoRabbitMQ(message []byte, config rabbitConfig){
	url := "amqp://" + config.User + ":" + config.Pass + "@" + config.Host + ":" + strconv.Itoa(config.Port) + "/" + config.Vhost
	conn, err := amqp.Dial(url)
	failOnError(err, "Failed to connect to RabbitMQ")
	defer conn.Close()

	ch, err := conn.Channel()
	failOnError(err, "Failed to open a channel")
	defer ch.Close()

	q, err := ch.QueueDeclare(
		config.Queue,
		true,
		false,
		false,
		false,
		nil,
	)
	failOnError(err, "Failed to declare a queue")

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
	failOnError(err, "Failed to publish a message")
}


func updateSender(config rabbitConfig) {
	messageJson := makeMessage()
	sendtoRabbitMQ(messageJson, config)
}
