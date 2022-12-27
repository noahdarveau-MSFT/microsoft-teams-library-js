import * as microsoftTeams from '@microsoft/teams-js';
import { GetServerSideProps } from 'next';
import Head from 'next/head';
import React, { ReactElement, useEffect, useState } from 'react';

export interface SSRProps {
  renderString: string;
  time: string;
}

export default function IndexPage(props: SSRProps): ReactElement {
  const [teamsContext, setTeamsContext] = useState({});

  useEffect(() => {
    microsoftTeams.app.initialize().then(() => {
      microsoftTeams.app.getContext().then((ctx) => {
        setTeamsContext(ctx);
      });
      microsoftTeams.app.notifySuccess();
    });
  }, []);

  return (
    <div>
      <Head>
        <title>SSR Test App</title>
      </Head>
      <div>
        <h1 id="id01">{props.renderString}</h1>
        <h1 id="time">The current server time is {props.time}</h1>
        <pre>
          <b>Context:</b> {JSON.stringify(teamsContext, null, 2)}
        </pre>
      </div>
    </div>
  );
}

/**
 * @returns prop data
 */
export const getServerSideProps: GetServerSideProps = async () => {
  const time = JSON.stringify({ time: new Date() });
  return {
    props: {
      renderString: 'This string brought to you by the server',
      time,
    },
  };
};