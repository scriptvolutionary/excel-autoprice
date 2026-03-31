class SeasonalPriceError(Exception):
    """Базовая ошибка приложения."""


class ValidationError(SeasonalPriceError):
    """Ошибка валидации входных данных."""


class DuplicateResolutionError(SeasonalPriceError):
    """Ошибка разрешения дублей клиентских файлов."""


class UnknownSkuError(SeasonalPriceError):
    """Ошибка сопоставления SKU."""


class PdfExportError(SeasonalPriceError):
    """Ошибка экспорта подтверждений в PDF."""
