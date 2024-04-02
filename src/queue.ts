import { Queue, Worker } from 'bullmq';
import IORedis from 'ioredis';
import { fetchAndSendEmailGoogle, fetchAndSendEmailOutlook } from "./task";
import { google } from 'googleapis';
import { Client } from '@microsoft/microsoft-graph-client';


const redisConnection = new IORedis({
  host: 'localhost',
  port: 6379,
  maxRetriesPerRequest: null
});
const emailQueue = new Queue('emailQueue', { connection: redisConnection });


const worker = new Worker('emailQueue', async (job) => {
  const { provider } = job.data;
  if (provider === "google") {
      const { accessToken } = job.data;
      const oauth2Client = new google.auth.OAuth2();
      oauth2Client.setCredentials({ access_token: accessToken });
      if (oauth2Client) {
          const res = await fetchAndSendEmailGoogle(oauth2Client);
          console.log(res);
      }
  } else {
      const { accessToken } = job.data;
      const graph = Client.init({
          authProvider: (done) => {
              done(null, accessToken);
          },
      });
      const res = await fetchAndSendEmailOutlook(graph);
      console.log(res);
  }
}, {
  connection: redisConnection
});

// Event listener for when the worker completes a job
worker.on('completed', (job) => {
  console.log(`Job ${job.id} has completed successfully`);
});

// Event listener for when the worker encounters an error
worker.on('failed', (job, err) => {
  console.error(`Job ${job.id} has failed with error: ${err}`);
});


export { emailQueue }