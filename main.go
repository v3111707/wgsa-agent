package main

import (
	"fmt"
	"log"
	"os"
	"strings"
	"golang.org/x/sys/windows/svc"
	"golang.org/x/sys/windows/svc/eventlog"
)

const svcName = "wgsa-agent"

func failOnError(err error, msg string) {
	if err != nil {
		log.Panicf("%s: %s", msg, err)
	}
}

func main() {
	elog, _ = eventlog.Open(svcName)
	var cmd string
	var err error
	defer elog.Close()

	if len(os.Args) > 1 {
		cmd = strings.ToLower(os.Args[1])
		elog.Info(1, fmt.Sprintf("Args[1]: %s ", os.Args[1]))
	}

	if len(os.Args) < 2 {
		usage("no command specified")
	}

	switch cmd {
	case "debug":
		runService(svcName, true)
		return
	case "install":
		if len(os.Args) > 2 {
			elog.Info(1, fmt.Sprintf("Args[2] %s ", os.Args[2]))
			err = installService(svcName, svcName, os.Args[2])
		} else {
			fmt.Println("Set path to config")
		}
	case "remove":
		err = removeService(svcName)
	case "start":
		err = startService(svcName)
	case "stop":
		err = controlService(svcName, svc.Stop, svc.Stopped)
	case "pause":
		err = controlService(svcName, svc.Pause, svc.Paused)
	case "continue":
		err = controlService(svcName, svc.Continue, svc.Running)
	case "gu":
		updates, err := getUpdatesList("IsInstalled=0")
		if err != nil {
			fmt.Println(err)
			elog.Info(1, fmt.Sprintf("Error: %s ", err))

		}
		fmt.Println(updates)
	default:
		isIntSess, err := svc.IsAnInteractiveSession()
		failOnError(err, "Failed to determine if we are running in an interactive session")

		if !isIntSess{
			runService(svcName, false)
			return
		}else {
			usage(fmt.Sprintf("invalid command %s", cmd))
		}
	}
	failOnError(err, "Failed to " + cmd )
	return
}