FROM php:8.3-apache

WORKDIR /var/www/html

# Copy application files
COPY . /var/www/html/

# Copy Apache configuration (must use port 80)
COPY apache.conf /etc/apache2/sites-available/000-default.conf

# Enable mod_rewrite
RUN a2enmod rewrite

# DO NOT run a2ensite (already enabled by default)
# RUN a2ensite 000-default.conf  <-- remove this

# Install required extensions for PhpSpreadsheet
RUN apt-get update && apt-get install -y \
    libzip-dev \
    unzip \
    && docker-php-ext-install zip pdo pdo_mysql \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set ServerName
RUN echo "ServerName api-spreadsheet.goodlifemicrolending.com" >> /etc/apache2/apache2.conf

ENV APP_URL=api-spreadsheet.goodlifemicrolending.com

# Railway expects container to listen on 80
EXPOSE 80

CMD ["apache2-foreground"]
