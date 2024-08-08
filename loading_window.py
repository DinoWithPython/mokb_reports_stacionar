"""Для отображения предзагрузки."""
import pyi_splash
# Игнорируем установку пакета, он уже входит в модуль pyinstaller

pyi_splash.update_text('loaded...')   
# Подробнее: https://pyinstaller.readthedocs.io/en/stable/advanced-topics.html#Module-pyi_splash

pyi_splash.close()