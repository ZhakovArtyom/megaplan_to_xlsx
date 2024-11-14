**Запуск приложения из корневой директории проекта**
1. Соберите Docker-образ:

`docker build -t megaplan-project-to-xlsx .`

2. Запустите Docker-контейнер:

_для linux:_
`docker run -d --name megaplan-container -v $(pwd)/logs:/app/logs -p 80:80 megaplan-project-to-xlsx`

_для windows:_
`docker run -d --name megaplan-container -v ${PWD}/logs:/app/logs -p 80:80 megaplan-project-to-xlsx`

**_Если нужно удалить контейнер для перезапуска кода:_**
`docker rm -f megaplan-container`