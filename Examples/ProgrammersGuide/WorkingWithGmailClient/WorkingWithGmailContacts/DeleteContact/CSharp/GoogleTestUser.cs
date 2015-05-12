using System;
using System.Collections.Generic;
using System.Text;

namespace Aspose.Email.Tests
{
    public class GoogleTestUser : TestUser
    {
        public GoogleTestUser(
            string name,
            string eMail,
            string password)
            : this(
                name,
                eMail,
                password,
                null,
                null,
                null)
        {
        }

        public GoogleTestUser(
            string name,
            string eMail,
            string password,
            string clientId,
            string clientSecret)
            : this(
                name,
                eMail,
                password,
                clientId,
                clientSecret,
                null)
        {
        }

        public GoogleTestUser(
            string name,
            string eMail,
            string password,
            string clientId,
            string clientSecret,
            string refreshToken)
            : base(
                name,
                eMail,
                password,
                "gmail.com")
        {
            ClientId = clientId;
            ClientSecret = clientSecret;
            RefreshToken = refreshToken;
        }

        public readonly string ClientId;
        public readonly string ClientSecret;
        public readonly string RefreshToken;
    }
}