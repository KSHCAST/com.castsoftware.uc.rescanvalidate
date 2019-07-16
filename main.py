#import cast_upgrade_1_5_21  #@UnusedImport
import os
import time
import xlsxwriter
from cast.application import ApplicationLevelExtension
from tkinter.font import BOLD


class Report(ApplicationLevelExtension):
    
    def start_application(self, application):
        """
        Called before analysis.
        
        .. versionadded:: CAIP 8.3
        
        :type application: :class:`cast.application.Application`
        @type application: cast.application.Application
        """
        pass
    
    
    def end_application(self, application):
        """
        Called at the end of application's analysis.

        :type application: :class:`cast.application.Application`
        @type application: cast.application.Application
        """
        pass

    def after_module(self, application):
        """
        Called after module content creation.
        
        .. versionadded:: CAIP 8.3
        
        :type application: :class:`cast.application.Application`
        @type application: cast.application.Application
        """
        
        pass
        

    def after_snapshot(self, application):
        """
        Called after module content creation.
        Gives you the central's application.
        
        .. versionadded:: CAIP 8.3
        
        :type application: :class:`cast.application.central.Application`        
        @type application: cast.application.central.Application
        """
        
        # do your things afte rmodule...
        # this import may fail in versions < 8.3
        from cast.application import publish_report # @UnresolvedImport
        
        # generate a path in LISA my_report<timestamp>.xlxs 
        report_path = os.path.join(self.get_plugin().intermediate, time.strftime("rescan_report%Y%m%d_%H%M%S.xlsx"))
        
        # calculate and fill file content
        
        # we use XlsxWriter to write excel 
        # @see https://xlsxwriter.readthedocs.io/
        workbook = xlsxwriter.Workbook(report_path)
        # @type workbook: xlsxwriter.Workbook
        
        worksheet = workbook.add_worksheet('Rescan Validations')
        # @type worksheet: xlsxwriter.Worksheet
        
        
        # kb represent a local
        #kb = application.get_knowledge_base()
        kb = application.get_application_configuration().get_analysis_service()
        central = application.get_central() 
        # @type kb : cast.application.KnowledgeBase
        aop_base_path = os.getcwd() + os.sep + 'RescanValidationChecks'
        print (aop_base_path)
        # we count elements in table keys
        aop_exe_name = 'RescanValidations.exe'
        schema_dtls = {}
        for sch_dtl in kb.execute_query("""SELECT CURRENT_USER username, inet_server_addr() host ,inet_server_port() port;"""):
            schema_dtls['username'] =  sch_dtl[0]
            schema_dtls['host'] =  sch_dtl[1].compressed
            schema_dtls['port'] =  sch_dtl[2]
            schema_dtls['password'] =  'CastAIP'
            schema_dtls['schema_name'] =  kb.name.rstrip('_local')
        
        print (schema_dtls)
        
        # calculate a status
        # this is a sample : choose one
        os.chdir(aop_base_path)
        ret = os.system(aop_exe_name + ' ' + str(schema_dtls['host']) + ' ' + str(schema_dtls['port']) + ' ' + str(schema_dtls['username']) + ' ' + str(schema_dtls['password']) + ' ' + str(schema_dtls['schema_name']))
        #ret = os.system(aop_exe_name + ' ' + schema_dtls['host'] +  ' ' +  schema_dtls['port'] + ' ' +  schema_dtls['username'] +  ' ' +   schema_dtls['password'] + ' ' + schema_dtls['schema_name'] )
        print (str(ret))
        #status = "Warning"
        #status = "KO"
        status = "Warning"
        
        if ret == 0:
            status = "OK"
            
         
        worksheet.write(0, 0, 'Final Verdict')
        worksheet.write(0, 1, status)
        workbook.close()
        # publish report :  
        publish_report('Rescan Validations', 
                       status, "Rescan completeness", '', detail_report_path=report_path)
