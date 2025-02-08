package tw.topgiga.zygoreport.processFactories;

import org.adempiere.base.IProcessFactory;
import org.compiere.process.ProcessCall;
import org.osgi.service.component.annotations.Component;

import tw.topgiga.zygoreport.report.ZYGOReportProcess;

@Component(

		property = { "service.ranking:Integer=2" }, service = org.adempiere.base.IProcessFactory.class)
public class ZYGOReportProcessFactory implements IProcessFactory {

	@Override
	public ProcessCall newProcessInstance(String className) {
		// TODO Auto-generated method stub
		if (className.equals(ZYGOReportProcess.class.getName())) {
			return new ZYGOReportProcess();
		}
		return null;
	}

}
