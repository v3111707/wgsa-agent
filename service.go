package main

import (
	"fmt"
	"time"
	"os"
	"golang.org/x/sys/windows/svc"
	"path/filepath"
	"golang.org/x/sys/windows/svc/mgr"
	"golang.org/x/sys/windows/svc/eventlog"
	"golang.org/x/sys/windows/svc/debug"
	"github.com/BurntSushi/toml"
)

var elog debug.Log

type tomlConfig struct {
	Main mainConfig
	RabbitMQ rabbitConfig
}

type mainConfig struct {
	TimePeriod int
}

type rabbitConfig struct {
	Host string
	Port int
	Vhost string
	Queue string
	User string
	Pass string
}

type wgsaservice struct{}

func usage(errmsg string) {
	fmt.Fprintf(os.Stderr,
		"%s\n\n"+
			"usage: %s <command>\n"+
			"       where <command> is one of\n"+
			"       install, remove, debug, start, stop, pause or continue.\n",
		errmsg, os.Args[0])
	os.Exit(2)
}

func startService(name string) error {
	m, err := mgr.Connect()
	failOnError(err, "Failed ")
	defer m.Disconnect()
	s, err := m.OpenService(name)
	failOnError(err, "could not access service ")
	defer s.Close()
	err = s.Start("is", "manual-started")
	failOnError(err, "could not start service ")
	return nil
}

func controlService(name string, c svc.Cmd, to svc.State) error {
	m, err := mgr.Connect()
	failOnError(err, "Can't connect ")
	defer m.Disconnect()
	s, err := m.OpenService(name)
	failOnError(err, "could not access service: ")
	defer s.Close()
	status, err := s.Control(c)
	failOnError(err, "could not send control: ")
	timeout := time.Now().Add(10 * time.Second)
	for status.State != to {
		if timeout.Before(time.Now()) {
			return fmt.Errorf("timeout waiting for service to go to state=%d", to)
		}
		time.Sleep(300 * time.Millisecond)
		status, err = s.Query()
		if err != nil {
			return fmt.Errorf("could not retrieve service status: %v", err)
		}
	}
	return nil
}

func (m *wgsaservice) Execute(args []string, r <-chan svc.ChangeRequest, changes chan<- svc.Status) (ssec bool, errno uint32) {
	var config  tomlConfig
	configPath := os.Args[1]
	elog.Info(1, fmt.Sprintf("Using %s ", configPath))
	if _, err := toml.DecodeFile(configPath, &config); err != nil {
		fmt.Println(err)
		return
	}
	const cmdsAccepted = svc.AcceptStop | svc.AcceptShutdown | svc.AcceptPauseAndContinue
	changes <- svc.Status{State: svc.StartPending}
	fasttick := time.Tick(500 * time.Millisecond)
	slowtick := time.Tick(2 * time.Second)
	tick := fasttick
	changes <- svc.Status{State: svc.Running, Accepts: cmdsAccepted}
	updateSenderNextRunTime := time.Now()
loop:
	for {
		select {
		case <-tick:
			if(updateSenderNextRunTime.Before(time.Now())){
				elog.Info(1, "Run updateSender")
				go updateSender(config.RabbitMQ)
				updateSenderNextRunTime = time.Now().Add(time.Duration(config.Main.TimePeriod) * time.Minute)
			}
		case c := <-r:
			switch c.Cmd {
			case svc.Interrogate:
				changes <- c.CurrentStatus
				// Testing deadlock from https://code.google.com/p/winsvc/issues/detail?id=4
				time.Sleep(100 * time.Millisecond)
				changes <- c.CurrentStatus
			case svc.Stop, svc.Shutdown:
				break loop
			case svc.Pause:
				changes <- svc.Status{State: svc.Paused, Accepts: cmdsAccepted}
				tick = slowtick
			case svc.Continue:
				changes <- svc.Status{State: svc.Running, Accepts: cmdsAccepted}
				tick = fasttick
			default:
				elog.Error(1, fmt.Sprintf("unexpected control request #%d", c))
			}
		}
	}
	changes <- svc.Status{State: svc.StopPending}
	return
}

func runService(name string, isDebug bool) {
	var err error
	if isDebug {
		elog = debug.New(name)
	} else {
		elog, err = eventlog.Open(name)
		if err != nil {
			return
		}
	}
	defer elog.Close()

	elog.Info(1, fmt.Sprintf("starting %s service", name))
	run := svc.Run
	if isDebug {
		run = debug.Run
	}
	err = run(name, &wgsaservice{})
	if err != nil {
		elog.Error(1, fmt.Sprintf("%s service failed: %v", name, err))
		return
	}
	elog.Info(1, fmt.Sprintf("%s service stopped", name))
}

func exePath() (string, error) {
	prog := os.Args[0]
	p, err := filepath.Abs(prog)
	if err != nil {
		return "", err
	}
	fi, err := os.Stat(p)
	if err == nil {
		if !fi.Mode().IsDir() {
			return p, nil
		}
		err = fmt.Errorf("%s is directory", p)
	}
	if filepath.Ext(p) == "" {
		p += ".exe"
		fi, err := os.Stat(p)
		if err == nil {
			if !fi.Mode().IsDir() {
				return p, nil
			}
			err = fmt.Errorf("%s is directory", p)
		}
	}
	return "", err
}

func installService(name, desc, configPath string) error {
	exepath, err := exePath()
	if err != nil {
		return err
	}
	m, err := mgr.Connect()
	if err != nil {
		return err
	}
	defer m.Disconnect()
	s, err := m.OpenService(name)
	if err == nil {
		s.Close()
		return fmt.Errorf("service %s already exists", name)
	}
	s, err = m.CreateService(name, exepath, mgr.Config{DisplayName: desc, StartType: 2}, configPath)
	if err != nil {
		return err
	}
	defer s.Close()
	err = eventlog.InstallAsEventCreate(name, eventlog.Error|eventlog.Warning|eventlog.Info)
	if err != nil {
		s.Delete()
		return fmt.Errorf("SetupEventLogSource() failed: %s", err)
	}
	fmt.Println("Service installed successfully")
	return nil
}

func removeService(name string) error {
	m, err := mgr.Connect()
	failOnError(err, "")
	defer m.Disconnect()
	s, err := m.OpenService(name)
	failOnError(err, "service is not installed")
	defer s.Close()
	err = s.Delete()
	failOnError(err, "")
	err = eventlog.Remove(name)
	failOnError(err, "RemoveEventLogSource() failed: ")
	fmt.Println("Service removed successfully")
	return nil
}

