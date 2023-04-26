import pandas as pd
from django.core.management.base import BaseCommand
from django.core.mail import send_mail
from myapp.models import Product

class Command(BaseCommand):
    help = 'Import data from Excel file into database'

    def add_arguments(self, parser):
        parser.add_argument('file_path', type=str, help='Path to the Excel file')

    def handle(self, *args, **kwargs):
        file_path = kwargs['file_path']
        try:
            # Read Excel file into a pandas DataFrame
            df = pd.read_excel(file_path)
            
            # Loop through each row of the DataFrame and create a new Product object
            imported_count = 0
            failed_count = 0
            for _, row in df.iterrows():
                try:
                    product = Product(
                        name=row['Name'],
                        price=row['Price'],
                        description=row['Description'],
                        category=row['Category']
                    )
                    product.save()
                    imported_count += 1
                except Exception as e:
                    self.stderr.write(str(e))
                    failed_count += 1
            
            # Send email with import summary
            subject = 'Import data summary'
            message = f'Imported {imported_count} records, {failed_count} records failed'
            from_email = 'admin@example.com'
            recipient_list = ['recipient@example.com']
            send_mail(subject, message, from_email, recipient_list)
            
            
            self.stdout.write(self.style.SUCCESS('Data imported successfully!'))
        except
