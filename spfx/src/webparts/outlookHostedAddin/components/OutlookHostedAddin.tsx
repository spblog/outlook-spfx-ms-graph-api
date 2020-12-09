import * as React from 'react';
import styles from './OutlookHostedAddin.module.scss';
import { IOutlookHostedAddinProps } from './IOutlookHostedAddinProps';
import { FC, useState } from 'react';
import { PrimaryButton } from 'office-ui-fabric-react';
import { AadHttpClient } from '@microsoft/sp-http';

export const OutlookHostedAddin: FC<IOutlookHostedAddinProps> = (props) => {
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);

  const saveToOneDrive = async () => {
    setSaving(true);
    setSaved(false);

    const client = await props.context.aadHttpClientFactory.getClient('af0331af-c6fc-4087-97c0-e18ba0bc5527');
    const tenantId = props.context.pageContext.aadInfo.tenantId;
    const mailId = Office.context.mailbox.convertToRestId(props.context.sdks.office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v1_0);

    await client.post(`http://localhost:7071/api/SaveMail/${tenantId}/${mailId}`, AadHttpClient.configurations.v1, {});

    setSaving(false);
    setSaved(true);
  };

  return (
    <div className={styles.outlookHostedAddin}>
      <PrimaryButton text={saving ? "Saving..." : "Save to OneDrive"} onClick={saveToOneDrive} disabled={saving} />
      {saved &&
        <div>Your email was saved!</div>}
    </div>
  );
};
