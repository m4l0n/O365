from datetime import datetime, timedelta
import requests
import re
import pickle
import os
from bs4 import BeautifulSoup
import logging
import utils
from getpass import getpass
from apscheduler.schedulers.asyncio import AsyncIOScheduler


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class TokenExpiredError(Exception):
    """
    An exception class that is raised when the refresh token or access token is expired.

    Attributes
    ----------
    message : str
      Error message string.

    Methods
    -------
    __str__:
      Overwrites str() to return error message string.
    """
    def __init__(self, message, *args, **kwargs):
        self.message = message
        super().__init__(self.message)

    def __str__(self):
        """
        Overwrites str() to return error message string.

        Returns
        -------
        self.message : Error message string
        """
        return self.message


class TokenInvalidError(Exception):
    """
    An exception class that is raised when the refresh token or access token is invalid.

    Attributes
    ----------
    message : str
      Error message string.

    Methods
    -------
    __str__:
      Overwrites str() to return error message string.
    """
    def __init__(self, message, *args, **kwargs):
        self.message = message
        super().__init__(self.message)

    def __str__(self):
        """
        Overwrites str() to return error message string.

        Returns
        -------
        self.message : Error message string
        """
        return self.message


class Token:
    def __init__(self, refresh_token=None, access_token=None, expires_in=None):
        self.refresh_token = refresh_token
        self.access_token = access_token
        self.expires_in = expires_in
        self.expiration_datetime = None
        self.refresh_token_expiry = None

    def is_expired(self):
        return datetime.now() > self.expiration_datetime if self.expiration_datetime is not None else True

    def find_expiration_datetime(self):
        return datetime.now() + timedelta(seconds = self.expires_in)


class Account:
    def __init__(self, scopes=None):
        if scopes is not None and isinstance(scopes, list):
            self.headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                              'Chrome/100.0.4896.127 Safari/537.36 Edg/100.0.1185.50',
                'Origin': 'https://developer.microsoft.com',
                'Referer': 'https://developer.microsoft.com/'
            }
            self.oauth2_token_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
            self.oauth2_redirect_url = 'https://developer.microsoft.com/en-us/graph/graph-explorer'
            self.graph_api = 'https://graph.microsoft.com/v1.0/me/'
            self.scopes = scopes
            self.client_id = 'de8bc8b5-d9f9-48b1-a8ad-b748da725064'
            self.token = self.load_token()
            if self.token.is_expired():
                self.refresh_token()
            self.scheduler = AsyncIOScheduler(timezone = "Asia/Kuala_Lumpur")
            self.scheduler.add_job(self.persist_refresh_token, "interval", hours=1)
            self.scheduler.start()
        else:
            # Must catch this exception
            logger.critical("The oauth scopes are not set!")
            raise ValueError("The scopes is not set. Please define the scopes.")

    def load_token(self):
        try:
            with open("oauth.cache", 'rb') as f:
                token = pickle.load(f)
                logger.debug("Token is loaded from cache file!")
        except (FileNotFoundError, pickle.UnpicklingError):
            logger.error("No oauth cache file found!")
            refresh_token = getpass(prompt = "Enter a refresh token: ")
            utils.clear_last_input()
            token = Token(refresh_token)
            self.refresh_token()
        finally:
            return token

    def save_token(self):
        logger.debug("Token is saved into cache file!")
        pickle.dump(self.token, open("oauth.cache", "wb"))

    def clear_cache(self):
        logger.debug("Token cache is removed!")
        os.remove("oauth.cache")

    def refresh_token(self):
        payload = {
            'client_id': self.client_id,
            'redirect_uri': self.oauth2_redirect_url,
            'scope': " ".join(self.scopes),
            'refresh_token': self.token.refresh_token,
            'grant_type': 'refresh_token'
        }
        auth = requests.post(self.oauth2_token_url, data = payload, headers = self.headers)
        if auth.status_code == 200:
            self.token.access_token = auth.json()['access_token']
            self.token.refresh_token = auth.json()['refresh_token']
            self.token.expires_in = auth.json()['expires_in']
            self.token.expiration_datetime = self.token.find_expiration_datetime()
            self.token.refresh_token_expiry = datetime.now() + timedelta(hours = 24)
            logger.info("Token refreshed!")
            self.save_token()
        elif auth.status_code == 400:
            invalid_codes = [9002313, 900144]
            regex = re.compile(r'\n.*')
            error_string = regex.sub("", auth.json()['error_description'])
            if auth.json()['error_codes'][0] in invalid_codes:
                # Must catch this exception
                logger.error("Refresh token is invalid!")
                raise TokenInvalidError(error_string)
            else:
                # Must catch this exception
                logger.error("Refresh Token Expired!")
                self.clear_cache()
                raise TokenExpiredError(error_string)

    def two_hour_schedule(self):
        try:
            if self.token.is_expired():
                raise TokenExpiredError("Access token has expired!")
            headers = self.headers
            headers['Authorization'] = f'Bearer {self.token.access_token}'
            events = requests.get(f'{self.graph_api}calendarview?startdatetime={datetime.utcnow() - timedelta(minutes = 1)}'
                                  f'&enddatetime={datetime.utcnow() + timedelta(hours = 2)}', headers = self.headers)
            if (events.status_code == 200 or '200'):
                if len(events.json()['value']) == 0:
                    logger.info("No meeting URL found!")
                    return "No link found"
                for event in events.json()['value']:
                    soup = BeautifulSoup(event['body']['content'], "lxml")
                    meeting_url = soup.find('a', class_ = "me-email-headline")['href']
                    diff = abs(datetime.strptime(event['start']['dateTime'][:-1], "%Y-%m-%dT%H:%M:%S.%f") - datetime.utcnow())
                    if (timedelta(seconds = 0) <= diff <= timedelta(seconds = 20)):
                        logger.info("Meeting URL Found!")
                        return meeting_url
                    else:
                        logger.info("No meeting URL found!")
                        return "No link found"
            elif (events.status_code == 401):
                logger.error("Access token has expired!")
                raise TokenExpiredError("Access token has expired!")
            # Must catch this exception
            elif (events.status_code == 400):
                logger.error("Access token is invalid!")
                raise TokenInvalidError("Access token is invalid!")
        except TokenExpiredError:
            logger.error("Access Token Expired!")
            self.refresh_token()
            return self.two_hour_schedule()

    def persist_refresh_token(self):
        logger.debug("Checking for refresh token expiry...")
        if abs(self.token.refresh_token_expiry - datetime.now()) < timedelta(hours = 2):
            logger.info("Refresh token is expiring. Will be refreshed now.")
            self.refresh_token()
        else:
            logger.debug("Refresh token is still new!")


if __name__ == "__main__":
    scopes = [
        "Calendars.Read",
        "Calendars.Read.Shared",
        "Channel.ReadBasic.All",
        "IMAP.AccessAsUser.All",
        "openid profile",
        "Team.ReadBasic.All",
        "User.Read email"
    ]
    account = Account(scopes = scopes)