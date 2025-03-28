#!/bin/bash

BUILD_TIMESTAMP="$(cat /opt/pandoc/.build_timestamp)"
export PANDOC_SERVICE_BUILD_TIMESTAMP=${BUILD_TIMESTAMP}

if [ ! -d /var/run/dbus ]; then
    mkdir -p /var/run/dbus
fi

if ! pgrep -x 'dbus-daemon' > /dev/null; then
    if [ -f /run/dbus/dbus.pid ]; then
        rm /run/dbus/dbus.pid
    fi
    dbus_session_bus_address_filename="/tmp/dbus_session_bus_address";
    dbus-daemon --system --fork --print-address > ${dbus_session_bus_address_filename};

    if [ $? -ne 0 ]; then
        echo "Failed to start dbus-daemon, exiting"
        exit 1
    fi

    BUS_ADDRESS=$(cat ${dbus_session_bus_address_filename});
    export DBUS_SESSION_BUS_ADDRESS=${BUS_ADDRESS};
fi

poetry run python -m app.PandocServiceApplication &

wait -n

exit $?
