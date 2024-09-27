import { createDirectLine, StyleOptions } from "botframework-webchat";
import { useMemo } from "react";
import { Store } from "redux";
import ReactWebChat from "botframework-webchat";

interface Props {
  token: string;
  store: Store;
}

function Chatbot(props: Props) {
  const { token, store } = props;

  const directLine = useMemo(() => createDirectLine({ token }), []);

  const styleOptions: StyleOptions = {
    hideUploadButton: true,
  };

  return token ? <ReactWebChat directLine={directLine} store={store} styleOptions={styleOptions} /> : <div />;
}

export default Chatbot;
