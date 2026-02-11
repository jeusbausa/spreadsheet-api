FROM php:8.3-apache

WORKDIR /var/www/html

COPY . /var/www/html/

COPY apache.conf /etc/apache2/sites-available/000-default.conf

RUN a2enmod rewrite

RUN apt-get update && apt-get install -y \
    libzip-dev \
    unzip \
    && docker-php-ext-install zip pdo pdo_mysql

RUN echo "ServerName api-spreadsheet.goodlifemicrolending.com" >> /etc/apache2/apache2.conf

ENV APP_ENV=production

EXPOSE 80

CMD ["apache2-foreground"]
