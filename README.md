
# Power Point slides to stories
You can use this util for easy creating instagram stories from MS Power Point slides

##### Dependecies:
- `MS PowerPoint`


## How it works
- Приложение получает путь к файлу:
	- ~~Через интерфейс~~
	- Консоль
	- ~~Перетягиванием файла `pptx` на `exe` файл~~
- Сканирует содержимое слайдов, определяет типы: картинка, гифка, видео
- Рендерит
	- Картинки – в `PNG`
	- `GIF` – в 15 секундые `mp4`
	- Video – в `mp4` ролики аналогичной длины
- Результаты сохраняются в папку с названием как у презентации по тому же пути

### Notes
- Для корректного рендеринга рекомендую закрывать окно MS PP перед стартом.
- Подготовьте свою презентацию. Перед рендерингом запустите просмотр презентации. Резутат после рендеринга будет аналогичным.
- Если видео не запускается с самого начала видео (открытия слайда), укажите это в настройках видео / анимации
- Если на слайде видео, то устанавливается минимальная длительность слайда в 1 секунду.
	- Т.е. итоговое видео будет длительностью не меньше 1 секунды или равное длительности самого длинного видео на слайде.
	- Если нужна иная длительность, задайте в настройках длительности показа слайда или в настройках видео.
- Длительность отображения слайда можно использовать и в других случаях
- Бывает, что не сработали советы выше, к примеру, у меня на первом слайде были GIF и короткое видео, я задал длительность показа 15 сек, но видео не проигрывалось даже до середины и зависало. Похоже это глюк.
 - Чтобы зациклить такое видео его можно было переделать в GIF.
 - Отдельно зарендерить видео нужной длительности.
 - Пересохранить видео (редактором, к примеру, на телефоне) и вставить заново.
- Чтобы исправить отдельные слайды рекомендую сделать копию презентации и удалить лишние слайды, чтобы получить лишь исправленный слайд. В дальнейшем планирую добавить удобный интерфейс для этого.

## 💖 Support my work

<table align="center" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><a href="https://ko-fi.com/nikmedoed"><img src="https://img.shields.io/badge/Ko--fi-donate-FF5E5B?logo=kofi" alt="Ko-fi" border="0"></a></td>
    <td><a href="https://boosty.to/nikmedoed/donate"><img src="https://img.shields.io/badge/Boosty-donate-FB400B?logo=boosty" alt="Boosty" border="0"></a></td>
    <td><a href="https://paypal.me/etonikmedoed"><img src="https://img.shields.io/badge/PayPal-donate-00457C?logo=paypal" alt="PayPal" border="0"></a></td>
    <td><a href="https://yoomoney.ru/to/4100119049495394"><img src="https://img.shields.io/badge/YooMoney-donate-8b3ffd?logo=yoomoney" alt="YooMoney" border="0"></a></td>
    <td><a href="https://github.com/nikmedoed#-support-my-work"><img src="https://img.shields.io/badge/Other-more-lightgrey?logo=github" alt="Other" border="0"></a></td>
  </tr>
</table>
