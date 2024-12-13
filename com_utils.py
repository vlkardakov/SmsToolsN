
import FreeSimpleGUI as sg
import folium
import re
import os
import webbrowser
from webview import create_window
import webview

def map(routes, labels=None):
    """
    routes - список маршрутов, где каждый маршрут это список ссылок
    labels - список подписей для точек (опционально)
    """
    def extract_coordinates(urls):
        coordinates = []
        for url in urls:
            match = re.search(r'll=(\d+\.\d+),(\d+\.\d+)', url)
            if match:
                lon, lat = match.groups()
                coordinates.append((float(lat), float(lon)))
        return coordinates

    def create_map(routes_coordinates, point_labels=None):
        # Берем координаты первой точки первого маршрута для центра карты
        center_point = routes_coordinates[0][0] if routes_coordinates[0] else (55.755826, 37.6173)
        m = folium.Map(location=center_point, zoom_start=12)

        # Цвета для разных маршрутов
        colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'darkblue', 'darkgreen']

        for route_idx, coordinates in enumerate(routes_coordinates):
            color = colors[route_idx % len(colors)]

            # Добавляем маркеры
            for i, (lat, lon) in enumerate(coordinates):
                # Определяем подпись для точки
                if point_labels and len(point_labels) > i:
                    label = point_labels[i]
                else:
                    label = f"Точка { i +1}"

                # Создаем круглый маркер с подписью
                folium.CircleMarker(
                    location=[lat, lon],
                    radius=8,
                    popup=label,
                    color=color,
                    fill=True,
                    fill_color=color
                ).add_to(m)

            # Создаем линию только если в маршруте больше одной точки
            if len(coordinates) > 1:
                folium.PolyLine(
                    coordinates,
                    weight=2,
                    color=color,
                    opacity=0.8
                ).add_to(m)

        # Сохраняем карту
        m.save('map.html')

    def show_map_window():
        webview.create_window('Карта', url=os.path.abspath('map.html'), width=800, height=600)
        webview.start()

    # Обработка входных данных
    if isinstance(routes[0], str):
        # Если передан один маршрут (список ссылок)
        routes = [routes]

    # Извлекаем координаты для каждого маршрута
    routes_coords = [extract_coordinates(route) for route in routes]

    # Создаем и показываем карту
    create_map(routes_coords, labels)
    show_map_window()

# Примеры использования:

# Один маршрут из нескольких точек с подписями
urls = [
    "http://m.maps.yandex.ru/?l=maps&ll=043.600066,56.479053&pt=043.600066,56.479053&z=13",
    "http://m.maps.yandex.ru/?l=maps&ll=043.610066,56.489053&pt=043.610066,56.489053&z=13",
    "http://m.maps.yandex.ru/?l=maps&ll=043.620066,56.499053&pt=043.620066,56.499053&z=13"
]
labels = ["Старт", "Промежуточная", "Финиш"]
map(urls, labels)

# Несколько маршрутов
routes = [
    # Маршрут 1 (три точки)
    [
        "http://m.maps.yandex.ru/?l=maps&ll=043.600066,56.479053&pt=043.600066,56.479053&z=13",
        "http://m.maps.yandex.ru/?l=maps&ll=043.610066,56.489053&pt=043.610066,56.489053&z=13",
        "http://m.maps.yandex.ru/?l=maps&ll=043.620066,56.499053&pt=043.620066,56.499053&z=13"
    ],
    # Маршрут 2 (одна точка)
    [
        "http://m.maps.yandex.ru/?l=maps&ll=043.630066,56.479053&pt=043.630066,56.479053&z=13"
    ]
]

labels = ["12","13","11","14"]
map(routes, labels)