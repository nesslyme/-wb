@echo off
chcp 65001 >nul
REM ===========================================================
REM  Запуск выгрузки отзывов Wildberries (Windows).
REM  Положите этот .bat рядом с wb_reviews.py и запустите двойным кликом.
REM  Скрипт сам поставит зависимость requests и откроет интерактивный ввод.
REM ===========================================================
title Выгрузка отзывов Wildberries
cd /d "%~dp0"

REM Ищем рабочий Python (py launcher или python из PATH)
where py >nul 2>nul && (set "PY=py") || (set "PY=python")

echo Проверяю зависимости (requests)...
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ОШИБКА] Не удалось установить зависимости. Проверьте, что Python и pip установлены.
    pause
    exit /b 1
)

echo.
%PY% wb_reviews.py %*

echo.
pause
