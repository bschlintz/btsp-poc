import React from 'react';
import { Stack, Text, FontWeights } from 'office-ui-fabric-react';
import FilesList from './FilesList';

const boldStyle = { root: { fontWeight: FontWeights.semibold } };

export const App: React.FunctionComponent = () => {
  return (
    <Stack
      horizontalAlign="center"
      verticalAlign="start"
      verticalFill
      styles={{
        root: {
          width: '960px',
          margin: '20px auto',
          textAlign: 'center',
          color: '#605e5c'
        }
      }}
      gap={15}
    >
      <img
        src="https://raw.githubusercontent.com/Microsoft/just/master/packages/just-stack-uifabric/template/src/components/fabric.png"
        alt="logo"
      />
      <Text variant="xxLarge" styles={boldStyle}>
        BTSP: Blob to SharePoint Concept
      </Text>
      <FilesList />
    </Stack>
  );
};
