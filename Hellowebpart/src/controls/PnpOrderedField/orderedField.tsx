import * as React from 'react';


export const orderedItem = (item:any, index:number): JSX.Element => {
	return (
		<span>
			<i className={"ms-Icon ms-Icon--" + item.iconName} style={{paddingRight:'4px'}}/>
			{item.text}
		</span>
	);
};
