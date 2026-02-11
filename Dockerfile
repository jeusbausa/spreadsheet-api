# Base image
FROM php:8.3-apache

# Set working directory
WORKDIR /var/www/html

# Install required system libraries
RUN apt-get update && apt-get install -y \
    libzip-dev \
    libpng-dev \
    libjpeg-dev \
    libfreetype6-dev \
    unzip \
    git \
    curl \
    && docker-php-ext-configure gd --with-freetype --with-jpeg \
    && docker-php-ext-install gd zip pdo pdo_mysql \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install Composer from official image
COPY --from=composer:latest /usr/bin/composer /usr/bin/composer

# Enable Apache mod_rewrite
RUN a2enmod rewrite

# Copy composer files first (for better caching)
COPY composer.json composer.lock ./

# Install PHP dependencies (production mode)
RUN composer install \
    --no-dev \
    --optimize-autoloader \
    --no-interaction \
    --prefer-dist

# Copy application files
COPY . .

# Copy Apache configuration
COPY apache.conf /etc/apache2/sites-available/000-default.conf

# Avoid Apache warning
RUN echo "ServerName api-spreadsheet.goodlifemicrolending.com" >> /etc/apache2/apache2.conf

# Set environment
ENV APP_ENV=production

# Expose Apache default port
EXPOSE 80

# Start Apache
CMD ["apache2-foreground"]
