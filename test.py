import requests


def download_pdf(url, filename):
    # Send a GET request to the URL
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Write the content of the response (PDF file) to a file
        with open(filename, 'wb') as file:
            file.write(response.content)
        print(f"PDF file downloaded successfully: {filename}")
    else:
        print(f"Failed to download PDF file. Status code: {response.status_code}")


# Example usage
pdf_url = "https://iea-nosso.sabacloud.com/content/socialtenantngx/NyKE2XMarcfq8DcjJbSJdg/1722403728/1/0024QWRhbSAgSGFza2V0dC5wZGY=V2VkIEp1bCAzMSAwMToyODozMyBFRFQgMjAyNA==/0140MDA4MHNVWEpUeHZ3bEYzM1pIRllOSHdPamZiYXhaY1RNVFpTOUxpV1hYTFdISzJwRmdQbHNVY0hVL0RseExtdXJhWCtxTUF6d3hRdWp5TVhaZVhhMDAxNkgxZXVvTHNMQkxHYXFmTVo=V2VkIEp1bCAzMSAwMToyODozMyBFRFQgMjAyNA==/eot/Adam%20%20Haskett.pdf"
output_filename = "downloaded_sample.pdf"

download_pdf(pdf_url, output_filename)
