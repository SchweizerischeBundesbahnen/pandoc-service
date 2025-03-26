import time

import docker


def test_container():
    client = docker.from_env()

    image, _ = client.images.build(path=".", tag="pandoc_service", buildargs={"APP_IMAGE_VERSION": "1.0.0"})
    container = client.containers.run(image=image, detach=True, name="pandoc_service", ports={"9082": 9082})
    time.sleep(5)
    logs = container.logs()
    container.stop()
    container.remove()

    assert logs == b"INFO:root:Pandoc service listening port: 9082\n"
