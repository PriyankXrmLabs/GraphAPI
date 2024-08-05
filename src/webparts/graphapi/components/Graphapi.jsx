import React, { useState, useEffect } from 'react';
import { MSGraphClient } from '@microsoft/sp-http';
import styles from './Graphapi.module.scss';

const Graphapi = (props) => {
  const [userInfo, setUserInfo] = useState(null);
  const [emails, setEmails] = useState([]);
  
  useEffect(() => {
    const fetchGraphData = async () => {
      try {
        const client = await props.context.msGraphClientFactory.getClient();
        const user = await client.api('/me').get();
        setUserInfo(user);

        const emailResponse = await client.api('/me/mailFolders/inbox/messages').top(5).get();
        setEmails(emailResponse.value);
        
      } catch (error) {
        console.error('Error fetching data from Graph API:', error);
      }
    };

    fetchGraphData();
  }, [props.context.msGraphClientFactory]);

  return (
    <section className={styles.graphapi}>
      <div className={styles.welcome}>Hello, Welcome to Graph API Demo web part</div>
      
      {userInfo && (
        <div className={styles.userInfo}>
          <h2>User Information</h2>
          <p><strong>Name:</strong> {userInfo.displayName}</p>
          <p><strong>Email:</strong> {userInfo.mail}</p>
        </div>
      )}
      
      {emails.length > 0 && (
        <div className={styles.emails}>
          <h2>Recent Emails</h2>
          <ul>
            {emails.map((email) => (
              <li key={email.id} className={styles.emailItem}>
                <p><strong>From:</strong> {email.from.emailAddress.name}</p>
                <p><strong>Subject:</strong> {email.subject}</p>
              </li>
            ))}
          </ul>
        </div>
      )}
      
      <button className={styles.fetchButton}>Fetch Data</button>
    </section>
  );
};

export default Graphapi;
